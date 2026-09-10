from dmponline import DMPonline

TOKEN = "fill in"
EMAIL = "fill in"

dmp_api = DMPonline(
    TOKEN,
    token_user=EMAIL
)

dmps = dmp_api.plan_statistics(
    params={"remove_tests": "false"}
)

print(dmps)