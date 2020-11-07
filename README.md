# Mark all messaged sent to "Deleted" folder as read

Make sure your macro security settings are set correctly:

for Outlook 2010 and up: File, Options, Trust Center, Trust Center Settings, Macro Security otherwise, you'll need to use selfcert.exe to sign your macros to test them which I highly recommended

Email will be marked read when it is moved to a folder "Deleted".

Place the code in `ThisOutlookSession` module, you must restart Outlook.
