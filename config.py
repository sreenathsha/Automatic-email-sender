import os

EXCEL_FILE_PATH = os.path.join(
    os.path.dirname(__file__),
    "excel_files",
    "initial dummy file (2).xlsx"
)

SUBJECT_TEMPLATE = "Form Reminder - Milestone: {milestone} | Item: {lineitem}"

EMAIL_TEMPLATE = """
<html>
<body>
<p>Hi there,</p>
<p>Please fill out the form for <b>{milestone}</b>, item <b>{lineitem}</b>:</p>
<p><a href="{form_link}">{form_text}</a></p>
<p>Thanks,<br>Your Automation Bot 🤖</p>
</body>
</html>
"""
