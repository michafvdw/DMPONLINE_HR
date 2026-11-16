import logging
import re
import requests
import json
import pandas as pd
from datetime import date, timedelta
from pandas.tseries.offsets import BDay
from pandas import json_normalize


class DMPonline:
    url = "https://hr.dmponline.dcc.ac.uk/api/"
    token = None

    def __init__(self, token, token_user=None, verify=True, authenticate_v1=False):
        self.token = token
        self.token_user = token_user
        self.verify = verify
        self.bearer_token = None

        # This export uses API v0 because it needs the complete plan text.
        # API v0 authenticates with the personal API token directly; it does
        # not use the email address or an API v1 bearer token.
        # Keep token_user available for compatibility and for optional v1 use.
        if authenticate_v1:
            if not token_user:
                raise ValueError(
                    "token_user (your DMPonline email address) is required "
                    "when authenticate_v1=True."
                )
            self.bearer_token = self.get_bearer_token(token_user)
        elif token_user:
            logging.info(
                "DMPonline email address supplied. API v0 will authenticate "
                "using the personal API token; API v1 authentication is skipped."
            )
        else:
            logging.warning(
                "No email address supplied. API v0 can still authenticate "
                "using the personal API token."
            )

    def get_bearer_token(self, token_user):
        """
        Exchange the DMPonline API code for a bearer token.

        This optional API v1 method sends the email and API token as JSON,
        matching the DMPonline authentication example. The HR export itself
        uses API v0 and does not need to call this method.
        """
        headers = {
            "Accept": "application/json",
            "Content-Type": "application/json",
        }
        params = {
            "grant_type": "authorization_code",
            "email": token_user,
            "code": self.token,
        }

        auth_url = f"{self.url}v1/authenticate"

        try:
            response = requests.post(
                auth_url,
                json=params,
                headers=headers,
                verify=self.verify,
                timeout=30,
            )
        except requests.exceptions.RequestException as exc:
            raise RuntimeError(
                f"Could not connect to the DMPonline authentication endpoint: {exc}"
            ) from exc

        logging.info(
            "DMPonline authentication returned HTTP status %s.",
            response.status_code,
        )

        try:
            response_data = response.json()
        except ValueError as exc:
            raise RuntimeError(
                "DMPonline returned a response that was not valid JSON. "
                f"HTTP status: {response.status_code}. "
                f"Response: {response.text[:1000]}"
            ) from exc

        if not response.ok:
            # Do not print the submitted API token.
            raise RuntimeError(
                "DMPonline authentication failed. "
                f"HTTP status: {response.status_code}. "
                f"Response: {response_data}"
            )

        access_token = response_data.get("access_token")
        if not access_token:
            raise RuntimeError(
                "DMPonline responded, but the response did not contain "
                "'access_token'. This may indicate that the authentication "
                "endpoint, credentials, or request format has changed. "
                f"Response: {response_data}"
            )

        logging.info("Successfully authenticated with DMPonline.")
        return access_token

    def _last_businessday(self):
        """Return the previous business day as YYYY-MM-DD."""
        last_bday = date.today() - BDay(1)
        return last_bday.strftime("%Y-%m-%d")

    def _yesterday(self):
        """Return yesterday as YYYY-MM-DD."""
        yesterday = date.today() - timedelta(days=1)
        return yesterday.strftime("%Y-%m-%d")

    def get(self, request="v0/statistics/plans", params=None):
        """
        Make a GET request to the DMPonline API.

        Returns parsed JSON for HTTP 200. Returns None for HTTP/API errors.
        """
        if params is None:
            params = {"remove_tests": "true"}

        url = self.url + request
        headers = {
            "Authorization": f"Token token={self.token}",
            "Accept": "application/json",
            "Content-Type": "application/json",
        }

        if "api/v1" in url or "/v1/" in url:
            if not self.bearer_token:
                raise RuntimeError(
                    "This API v1 request requires a bearer token. "
                    "Provide token_user when creating DMPonline."
                )
            headers["Authorization"] = f"Bearer {self.bearer_token}"

        try:
            response = requests.get(
                url,
                headers=headers,
                params=params,
                verify=self.verify,
                timeout=60,
            )
        except requests.exceptions.RequestException as exc:
            logging.error("Request to %s failed: %s", url, exc)
            return None

        if response.status_code != 200:
            logging.error(
                "Request to %s returned HTTP status %s.\n%s",
                url,
                response.status_code,
                response.text[:2000],
            )
            return None

        try:
            return response.json()
        except ValueError:
            logging.error(
                "Request to %s returned invalid JSON:\n%s",
                url,
                response.text[:2000],
            )
            return None

    @staticmethod
    def process_users(x, key="email"):
        return "; ".join(
            [
                item[key]
                for item in json.loads(str(x).replace("'", '"'))
                if key in item
            ]
        )

    def get_plan(self, plan_id):
        """Retrieve a plan from the API."""
        try:
            # The API v0 endpoint may fail for some plans.
            df = self.get_plan_v0(plan_id)
        except requests.exceptions.ConnectionError:
            # Fall back to API v1.
            df = self.get_plan_v1(plan_id)

        return df

    def get_plan_v0(self, plan_id):
        """Retrieve plan metadata from API v0 as a DataFrame."""
        request = "v0/plans"
        data = self.get(request=request, params={
            "plan": plan_id,
            "remove_tests": "false",
        })

        if data == [] or data is None:
            return None

        df = json_normalize(data)
        df.creation_date = pd.to_datetime(df.creation_date)
        df.last_updated = pd.to_datetime(df.last_updated)
        df.rename(
            columns={"principal_investigator.email": "email_pi"},
            inplace=True,
        )
        df.users = df.users.apply(
            lambda x: self.process_users(x, key="email")
        )

        return df

    def get_plan_v1(self, plan_id):
        """Retrieve plan metadata from API v1 as a DataFrame."""
        request = f"v1/plans/{plan_id}"
        data = self.get(
            request=request,
            params={"remove_tests": "false"},
        )

        if not data or not data.get("items"):
            return None

        df = json_normalize(data["items"][0]["dmp"])
        df.created = pd.to_datetime(df.created)
        df.modified = pd.to_datetime(df.modified)

        if "contributor" in df:
            try:
                logging.debug("Reading contributors.")
                df.contributor = df.contributor.apply(
                    lambda x: self.process_users(x, key="mbox")
                )
            except (ValueError, TypeError, KeyError):
                logging.debug(
                    "Could not parse contributor data as JSON; "
                    "extracting email addresses instead."
                )

                def process_contributor(x):
                    return "; ".join(
                        re.findall(r"[a-z.-]+@[a-z.-]+", str(x), re.I)
                    )

                df.contributor = df.contributor.apply(process_contributor)

        return df

    def dmp_count(self):
        """Return the number of DMPs."""
        data = self.get()
        if not data or "plans" not in data:
            return 0

        df = json_normalize(data["plans"])
        logging.debug("Number of plans: %s", df.shape[0])
        return df.shape[0]

    def plan_statistics(self, params=None):
        """
        Return metadata about all plans in the organisation as a DataFrame.
        """
        if params is None:
            params = {"remove_tests": "true"}

        data = self.get(
            request="v0/statistics/plans",
            params=params,
        )

        if data is None or "plans" not in data:
            return None

        df = json_normalize(data["plans"])
        df.date_created = pd.to_datetime(df.date_created)
        df.date_last_updated = pd.to_datetime(df.date_last_updated)
        return df

    def get_departments(self):
        """Return the organisation's departments as a DataFrame."""
        data = self.get(request="v0/departments", params={})
        if data is None:
            return None

        return pd.json_normalize(data)

    def get_department_users(self):
        """Return departments and their associated users as a DataFrame."""
        data = self.get(request="v0/departments/users", params={})
        if data is None:
            return None

        df_list = []
        for _, row in pd.json_normalize(data).iterrows():
            df_fac = pd.DataFrame(row.users)
            df_fac["faculty_en"] = row.code
            df_list.append(df_fac)

        if not df_list:
            return pd.DataFrame()

        return pd.concat(df_list, ignore_index=True)

    def has_personal_data(self, plan_id, verbose=False):
        """
        Determine whether personal data is associated with a DMP.

        Returns True, False, or None if the template is unsupported or the
        relevant question cannot be found/read.
        """
        parsed = self.get(
            "v0/plans",
            params={"plan": plan_id, "remove_tests": "false"},
        )

        if not parsed:
            logging.warning("No plan data returned for plan_id %s.", plan_id)
            return None

        try:
            template_id = parsed[0]["template"]["id"]
            plan_content = parsed[0]["plan_content"][0]
        except (IndexError, KeyError, TypeError):
            logging.warning(
                "Could not read template/plan content for plan_id %s.",
                plan_id,
            )
            return None

        template_question_mapping = {
            975303870: (5, 2),       # TU Delft Data Management Plan (2021)
            1753695087: (5, 1),      # Data Management Plan NWO (September 2020)
            1506827492: (5, 1),      # NWO Data Management Plan (January 2020)
            1461074155: (2, 1),      # Data management ZonMw-template 2019
            1165855271: (1, 10),     # TU Delft Data Management Questions
        }

        if template_id not in template_question_mapping:
            logging.warning(
                "Plan %s uses unsupported template_id %s.",
                plan_id,
                template_id,
            )
            return None

        section_number, question_number = template_question_mapping[
            template_id
        ]

        target_question = None

        for section in plan_content.get("sections", []):
            if section.get("number") == section_number:
                for question in section.get("questions", []):
                    if question.get("number") == question_number:
                        target_question = question
                        break
                break

        if target_question is None:
            logging.warning(
                "Could not find the personal-data question for plan %s.",
                plan_id,
            )
            return None

        if not target_question.get("answered"):
            return None

        answer = target_question.get("answer") or {}
        options = answer.get("options", []) if isinstance(answer, dict) else []

        if not options or not isinstance(options[0], dict):
            return None

        personal_data_str = options[0].get("text")

        if verbose:
            logging.info("Template: %s", parsed[0]["template"].get("title"))
            logging.info("Question: %s", target_question.get("text"))
            logging.info("Answer: %s", personal_data_str)

        if personal_data_str == "Yes":
            return True
        if personal_data_str == "No":
            return False

        return None
