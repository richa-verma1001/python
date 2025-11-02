# Example of using the OpenAI API with input from the user
# This example uses the OpenAI API to translate text into Hindi and Russian.
# This uses openai-2.6.1

from openai import OpenAI
client = OpenAI()

text = input("Enter a text for translation: ")
prompt = (f"Translate the text delimited by triple backticks into Russian```{text}```. ")

response = client.responses.create(
    model="gpt-5",
    input= prompt
)

print(response.output_text)