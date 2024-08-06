# Integrate Outlook with Todoist 

# Goal

Transform an e-mail in Microsoft Outlook into an Todoist task

# Integration

Todoist Rest API Documentation: https://developer.todoist.com/rest/v2

## Todoist API

### Get Todoist API Auth token

Get personal API token from the [account integrations settings](https://todoist.com/prefs/integrations).

```
export token="<PERSONAL_API_TOKEN>"
```

### Get Todoist projectId

```
curl -X GET \
  https://api.todoist.com/rest/v2/projects \
  -H "Authorization: Bearer $token"
```

E-mails will become tasks in the Inbox project, whose id is something like: `2326943403`.
* Note that the standard project is Inbox, thus its id is not needed to create a task inside it.

```
export projectId="2326943403"
```

### Create task for the given projectId

Use random uuid to avoid duplicated requests triggering executions twice.

```
export message="e-mail subject"
```

```
curl "https://api.todoist.com/rest/v2/tasks" \
  -X POST \
  --data "{\"content\": \"$message\", \"project_id\": \"$projectId\", \"labels\":[\"Amazon\"]}" \
  -H "Content-Type: application/json" \
  -H "X-Request-Id: $(uuidgen)" \
  -H "Authorization: Bearer $token"
```

## Outlook integration with Automator

Automator is a macOS tool for creating automated workflows. Using the AppleScript API, the following workflow adds the selected e-mail to the Todoist Inbox project and moves the e-mail to the folder “Todoist tasks” (remember to create it):

```
on run
    tell application "Microsoft Outlook"
        set selectedMessages to selected objects
        if selectedMessages is {} then
            display notification "No message selected. Select a message in Outlook before running the script"
        else
            set messageId to id of item 1 of selectedMessages
            set messageSubject to subject of item 1 of selectedMessages

            # Replace all occurences of "\&" for "and"
            set messageSubject to do shell script "echo " & quoted form of messageSubject & " | sed 's/\\&/and/g'"

            # Replace all occurences of "\'" for "\"
            set messageSubject to do shell script "echo " & quoted form of messageSubject & " | sed \"s/\\'//g\""

            set outlookUri to "outlook://" & messageId
            set the clipboard to outlookUri

            # Add a new Todoist task to the Inbox project by sending a POST request
            set todoistToken to "<PERSONAL_API_TOKEN>"
            set uuid to do shell script "uuidgen"
            set curl_task to do shell script ¬
                "curl \"https://api.todoist.com/rest/v2/tasks\" " &¬
                "-X POST " &¬
                "--data '{ " &¬
                    "\"content\": \"" & messageSubject & "\", " &¬
                    "\"description\": \"Outlook URL: " & outlookUri & "\", " &¬
                    "\"priority\": 4, " &¬
                    "\"due_string\": \"today\", " &¬
                    "\"labels\": [\"Amazon\"]}' " &¬
                "-H \"Content-Type: application/json\" " &¬
                "-H \"X-Request-Id: " & uuid & "\" " &¬
                "-H \"Authorization: Bearer " & todoistToken & "\""

            # Move e-mail to "Todoist tasks" e-mail folder
            move item 1 of selectedMessages to mail folder "Todoist tasks"

            display notification "Added task in Todoist: " & messageSubject
        end if
    end tell
 end run
```

This creates the following task:
[Image: Image.jpg]
## Open  custom outlook link

Follow section “Opening outlook links: a custom protocol handler” from the blog post: http://blog.hakanserce.com/post/outlook_automation_mac/


1. Create an AppleScript that implements a custom protocol handler to open location:

```
on open location this_URL
    set the messageId to text 11 thru -1 of this_URL
    tell application "Microsoft Outlook"
        activate
        open message id messageId
    end tell
end open location
```

1. Save the script as an app package in `Script Editor`. (`Save...` > `Format: Application`)
2. Open `Finder` and locate the saved file. Right click and select `Show Package Contents`
3. Under the Contents folder, open `Info.plist` and add the following content before the last two lines:

```
      <key>CFBundleIdentifier</key>
      <string>org.personal.outlook</string>
      <key>CFBundleURLTypes</key>
      <array>
          <dict>
              <key>CFBundleURLName</key>
              <string>Pass To OutlookUriHandler</string>
              <key>CFBundleURLSchemes</key>
              <array>
                  <string>outlook</string>
              </array>
          </dict>
      </array>
```

1. Save the file. Go back to your app and double click to execute it.

* I think that executing it once in life is enough
