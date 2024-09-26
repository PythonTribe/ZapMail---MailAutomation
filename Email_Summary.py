import os
from openai import OpenAI

# Initialize the OpenAI client with the API key
client = OpenAI(
    api_key=os.getenv("OPENAI_API_KEY"),
)

# Define the file path for the email content
email_file_path = r"C:\Users\Admin\PycharmProjects\Project1\Smart Email Application\email_output.txt"

# Read the email content from the file
with open(email_file_path, 'r', encoding='utf-8') as file:
    email_content = file.read()

# Create a completion request to summarize the email content
completion = client.chat.completions.create(
    model="gpt-3.5-turbo",
    messages=[
        {"role": "user", "content": f"Summarize the content: {email_content} pointwise in 3 new lines - "
                                    f"1.Subject 2.Meeting Agenda 3.Important dates.provide the email thread output in reverse order, most recent is most relevant.Summary relvant and dynamic"}
    ]
)

# Print the summary
print(completion.choices[0].message.content.strip())
