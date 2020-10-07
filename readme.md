# Energy Emailer

Energy Report Card emailer for normative behavior research at Seoul National University.

This application was developed for a colleague to assist in his research on the impact of social norms on people's energy usage habits. It takes as input an Excel spreadsheet containing email addresses, monthly energy usage, and flags indicating the type of report the user should receive, and then sends customized "report cards" to each address.

An example of the spreadsheets used as input can be found in `./ExcelFiles/`. The "email address" column contains my personal email account for all rows--please change it and refrain from bombing my inbox :)

## Rev 2

This is modified to fit the need of a new study in Unversity of Michigan.

Notes:

- The HTML template has to go through a css inliner to be a valid email document.
- Image hosting service may need change.
