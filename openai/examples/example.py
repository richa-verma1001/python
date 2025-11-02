# Example of using the OpenAI API
# This example uses the OpenAI API to generate a one-sentence bedtime story about a unicorn.

from openai import OpenAI
client = OpenAI()

response = client.responses.create(
    model="gpt-5",
    input="Write a one-sentence bedtime story about a unicorn."
)

print(response.output_text)