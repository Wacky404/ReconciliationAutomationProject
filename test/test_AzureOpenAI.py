from openai import AzureOpenAI
from azure.identity import DefaultAzureCredential, get_bearer_token_provider
from dotenv import load_dotenv

import unittest
import os

load_dotenv()
endpoint = os.getenv("ENDPOINT_URL")
deployment = os.getenv("DEPLOYMENT_NAME")
version = os.getenv("API_VERSION")
bearer = os.getenv("TOKEN_URL")


class TestAzureAI(unittest.TestCase):
    def test_gpt(self):
        token_provider = get_bearer_token_provider(
            DefaultAzureCredential(),
            bearer,
        )
        self.assertIsNotNone(token_provider, "Token was not provided!")

        client = AzureOpenAI(
            azure_endpoint=endpoint,
            azure_ad_token_provider=token_provider,
            api_version=version,
        )
        self.assertIsNotNone(client, "Client was not initiated!")

        completion = client.chat.completions.create(
            model=deployment,
            messages=[
                {
                    "role": "user",
                    "content": "What are the differences between Azure Machine Learning and Azure AI services?"
                }],
            max_tokens=800,
            temperature=0.7,
            top_p=0.95,
            frequency_penalty=0,
            presence_penalty=0,
            stop=None,
            stream=False
        )
        self.assertIsNotNone(completion, "The completion was not returned!")

        print(completion.to_json())

        self.assertTrue(True)


if __name__ == '__main__':
    unittest.main()
