# ProjectManagementReportingSoftware

Management reporting scripts. For each python script to work, one has to edit the Config file for the corresponding script with the necessary credentials.

## Project Management web services
[ClickUp](https://clickup.com/)
[TimeButler](https://timebutler.de/)
[Toggl](https://toggl.com/track/)

### The directory structure

```
├── README.md                         <- The README file for the users of this project.
├── ClickUp
│   ├── ClickUp_Config.json           <- The config file of the ClickUp script.
│   └── ClickUp_Tasks_pull_script.py  <- The ClickUp script.
├── TimeButler
│   ├── TButler_Config.json           <- The config file of the TimeButler script.
│   └── TButler.py                    <- The TimeButler script.
├── Toggl
    ├── Toggl_Config.json             <- The config file of the Toggl script.
    └── toggl_data_pull_script.py     <- The Toggl script.
```
### Pre-requisite

#### For all:

- Python3.7
-  pandas
- openpyxl

#### For ClickUp script:

  - pyclickup

#### For TimeButler script:

  - requests

#### For Toggl script:

  - togglPy (pip install -U TogglPy)
  - joblib
