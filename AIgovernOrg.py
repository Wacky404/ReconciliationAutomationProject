import openai
import pandas

# Deciding if dataframe should be turned into dictionary or left as df to be passed through into chatGPT
# to get best results. Researching.

# entire california database
california = pandas.read_csv(
    r"C:\Users\Wayne Cole\Downloads\Work Stuff\CSV California Educational Institutions 2023-06-20.csv"
)
print(california.head())
# information that is being passed to chatGPT
info = california[["GOVERNING_ORGANIZATION_NAME", "GOV_ADDRESS_LINE_1", "GOV_ADDRESS_LINE_2",
                  "GOV_MUNICIPALITY", "GOV_STATE_REGION_SHORT", "GOV_COUNTRY_CODE"]]
print(info.head())
# fields the data will be inputted into
collected_gov_phone = california["GOV_PhoneNumberFull"]
collected_gov_inactive = california["GOV_INACTIVE"]
collected_gov_dateclosed = california["GOV_INACTIVE_DATE"]
# How I will interaction will openai API, code will change.
API_KEY = open(r"C:\Users\Wayne Cole\Downloads\Work Stuff\API Key.txt").read()
openai.api_key = API_KEY

response = openai.ChatCompletion.create(
    model="gpt-3.5-turbo",
    message=[
        {"role": "user", "content": ""}
    ]
)
print(response)
