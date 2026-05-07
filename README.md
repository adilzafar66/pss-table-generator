# Table Generator

* [1\. Setup](#1.-Setup)
* [2\. Inputs & Outputs](#2.-Inputs-%26-Outputs)
* [3\. Naming Conventions](#3.-Naming-Conventions)
  * [3.1 Scenario Creation](#3.1-Scenario-Creation)
  * [3.2 Short Circuit Reports](#3.2-Short-Circuit-Reports)
  * [3.3 Device Duty Reports](#3.3-Device-Duty-Reports)
  * [3.4 Arc Flash Reports](#3.4-Arc-Flash-Reports)
* [4\. Short Circuit](#4.-Short-Circuit)
  * [4.1 Prerequisites](#4.1-Prerequisites)
  * [4.2 Operation](#4.2-Operation)
* [5\. Device Duty](#5.-Device-Duty)
  * [5.1 Prerequisites](#5.1-Prerequisites)
  * [5.2 Study Options](#5.2-Study-Options)
    * [5.2.1 Include Switches](#5.2.1-Include-Switches)
    * [5.2.2 Use All Switching Configs](#5.2.2-Use-All-Switching-Configs)
    * [5.2.3 Add Series Ratings](#5.2.3-Add-Series-Ratings)
    * [5.2.4 Mark Assumed Equipment](#5.2.4-Mark-Assumed-Equipment)
  * [5.3 Built-in Exclusions](#5.3-Built-in-Exclusions)
  * [5.4 Switchgear Calculation](#5.4-Switchgear-Calculation)
* [6\. Arc Flash](#6.-Arc-Flash)
  * [6.1 Prerequisites](#6.1-Prerequisites)
  * [6.2 Study Options](#6.2-Study-Options)
    * [6.2.1 Use SI Units](#6.2.1-Use-SI-Units)
    * [6.2.2 Incident Energy: High](#6.2.2-Incident-Energy%3A-High)
    * [6.2.3 Incident Energy: Critical](#6.2.3-Incident-Energy%3A-Critical)
    * [6.2.4 Include Revisions](#6.2.4-Include-Revisions)
  * [6.4 Output](#6.4-Output)
* [7\. Exclusions](#7.-Exclusions)
  * [7.1 Element ID Starting With](#7.1-Element-ID-Starting-With)
  * [7.2 Element ID Containing](#7.2-Element-ID-Containing)
* [8\. Menu Bar](#8.-Menu-Bar)
* [9\. Uninstallation](#9.-Uninstallation)
* [10\. References](#10.-References)

The PSS Table Generator is developed to automate the Excel table generation process for the PSS team. The program extracts data from ETAP via communicating with the ETAP API and parses that data to create Short-Circuit, Device Duty and Arc Flash tables for the inclusion in PSS Reports.

The program has been designed to work with minimal inputs and requires no external software or library to work besides a working installation of ETAP.

Open image-20241015-182911.png

![image-20241015-182911.png](https://media-cdn.atlassian.com/file/c1a00c7b-a7eb-4254-b511-1b74d07dfe18/image/cdn?allowAnimated=true&client=4ead6351-03ca-42fb-bbd7-8037c700c997&collection=contentId-3060957274&height=125&max-age=2592000&mode=full-fit&source=mediaCard&token=eyJhbGciOiJIUzI1NiJ9.eyJpc3MiOiI0ZWFkNjM1MS0wM2NhLTQyZmItYmJkNy04MDM3YzcwMGM5OTciLCJhY2Nlc3MiOnsidXJuOmZpbGVzdG9yZTpjb2xsZWN0aW9uOmNvbnRlbnRJZC0zMDYwOTU3Mjc0IjpbInJlYWQiXX0sImV4cCI6MTc3ODE4MTc0MiwibmJmIjoxNzc4MTc4ODYyLCJhYUlkIjoiNjEzNzlkMTUxNTZhYjMwMDcxNzVkZmNmIiwiaHR0cHM6Ly9pZC5hdGxhc3NpYW4uY29tL2FwcEFjY3JlZGl0ZWQiOmZhbHNlLCJhdXRoVHlwZSI6InNlc3Npb24ifQ.NKxaVnseEdbut9ORpLcNpszC1L2IB23pgyNioAB3B-g&width=624#media-blob-url=true&id=c1a00c7b-a7eb-4254-b511-1b74d07dfe18&clientId=4ead6351-03ca-42fb-bbd7-8037c700c997&contextId=contentId-3060957274&collection=contentId-3060957274)

Figure 1. User Interface of the PSS Table Generator

# 1\. Setup

The setup process is straightforward. It is recommended to install the application in the default location (Program Files x86), although the user can choose any location on their system. The installer will take a minute or two to generate all required files and will create a desktop shortcut for easy access.

The setup is located at: **B:\\Prime Documents\\PSS Projects\\PSS Tools\\PSS Table Generator**

# 2\. Inputs & Outputs

The application has a graphical interface currently appearing with 19 different inputs, with a minimum of 3 required inputs. The table below lists all the inputs of the program and provides a brief summary of their functionalities.

| Group | Input | Function |
| :-- | :-- | :-- |
| 1 | Reports | Short Circuit |
| 2 | Device Duty | Include device duty study type. |
| 3 | Arc Flash | Include arc flash study type. |
| 4 | Options | Create Scenarios |
| 5 | Run Scenarios | Run scenarios in ETAP for the selected study type. |
| 6 | Create Reports | Generate summary tables for the selected study types. |
| 7 | Device Duty | Include Switches |
| 8 | Use All Switching Configs | Use all switching configurations available in ETAP during scenario creation and execution. |
| 9 | Add Series Ratings | Add series rating values to series rated equipment in device duty tables from ETAP. |
| 10 | Mark Assumed Equipment | Add an asterisk to equipment with assumed ratings in device duty tables. |
| 11 | Arc Flash | Use SI Units |
| 12 | Incident Energy: High | Highlight equipment in orange having incident energy above this value in arc flash table. |
| 13 | Incident Energy: Critical | Highlight equipment in red having incident energy above this value in arc flash table. |
| 14 | Include Revisions | Revisions to include when creating arc flash scenarios in ETAP. |
| 15 | Exclude | Elements Starting With |
| 16 | Element Containing | Exclude elements that contain any of the following words in their names. |
| 17 | All / Except | Exclude all elements that obey any of the above two conditions / except the following. |
| 18 | Input Dir | ETAP Directory |
| 19 | Output Dir | Output Directory |
| 20 | Use ETAP Directory | Use ETAP project directory for output tables. |

# 3\. Naming Conventions

The functionality of the PSS Table Generator for automating report generation and scenario manipulation is highly dependent on a well-defined naming convention. This convention serves as a foundation, enabling the program to accurately identify and process files and information in a structured manner. Additionally, it helps establish a consistent workflow, ensuring efficiency and ease of use across different projects.

However, overly complex naming conventions can introduce errors and inefficiencies, potentially affecting the program’s performance. To mitigate this, the naming conventions for the PSS Table Generator have been kept as simple as possible, ensuring the program can locate the correct files and configurations without burdening the user.

The following sections provide detailed guidelines on the naming conventions required for each component of the application.

## 3.1 Scenario Creation

A key feature of the PSS Table Generator is its capability to automatically create and execute scenarios within ETAP. However, it is important to note that the scenarios are not generated entirely from scratch. Users must first define study cases, switching configurations, and revisions (if applicable). The program then permutates these variables to generate a scenario for each possible combination.

Adherence to a strict naming convention is crucial for the program to produce accurate scenarios. Each variable must be named according to the guidelines outlined in the table below to ensure correct scenario generation.

| Variable | Short Circuit | Device Duty | Arc Flash |
| :-- | :-- | :-- | :-- |
| Study Case | SC | SC and/orSC_IEC | AF_VCB orAF_VCBB |
| Switching Configuration | Any | Any | Any |
| Revision | N/A | N/A | Any |

## 3.2 Short Circuit Reports

There is no strict naming convention required for the short circuit table generation. The only condition is that all the short circuit reports output by ETAP must be named starting with `SC_`. Some examples are as follows:

1. SC\_PRES
2. SC\_ULT
3. SC\_Generator
  

Any report files contained in the ETAP project folder that do not start with the `SC_` will not be included in the short circuit table generation.

## 3.3 Device Duty Reports

To ensure the program runs smoothly and without any logical inconsistencies, users must follow the following naming convention when naming the Device Duty output reports in ETAP:

| Scenario | ANSI Report Name | IEC Report Name |
| :-- | :-- | :-- |
| Present | DD_PRES | DD_PRES_IEC |
| Ultimate | DD_ULT | DD_ULT_IEC |
| Generator | DD_GEN | DD_GEN_IEC |
| Other | DD_<COLUMNNAME> | DD_<COLUMNNAME>_IEC |

The output reports can be named while creating a Scenario. Refer to Figure 2 for an example.

Open image-20240719-221217.png

![image-20240719-221217.png](https://media-cdn.atlassian.com/file/580bb7f2-c0d5-4043-ab00-5395dd079810/image/cdn?allowAnimated=true&client=4ead6351-03ca-42fb-bbd7-8037c700c997&collection=contentId-3060957274&height=125&max-age=2592000&mode=full-fit&source=mediaCard&token=eyJhbGciOiJIUzI1NiJ9.eyJpc3MiOiI0ZWFkNjM1MS0wM2NhLTQyZmItYmJkNy04MDM3YzcwMGM5OTciLCJhY2Nlc3MiOnsidXJuOmZpbGVzdG9yZTpjb2xsZWN0aW9uOmNvbnRlbnRJZC0zMDYwOTU3Mjc0IjpbInJlYWQiXX0sImV4cCI6MTc3ODE4MTc0MiwibmJmIjoxNzc4MTc4ODYyLCJhYUlkIjoiNjEzNzlkMTUxNTZhYjMwMDcxNzVkZmNmIiwiaHR0cHM6Ly9pZC5hdGxhc3NpYW4uY29tL2FwcEFjY3JlZGl0ZWQiOmZhbHNlLCJhdXRoVHlwZSI6InNlc3Npb24ifQ.NKxaVnseEdbut9ORpLcNpszC1L2IB23pgyNioAB3B-g&width=563#media-blob-url=true&id=580bb7f2-c0d5-4043-ab00-5395dd079810&clientId=4ead6351-03ca-42fb-bbd7-8037c700c997&contextId=contentId-3060957274&collection=contentId-3060957274)

Figure 2. Setting Output Report for Device Duty scenario in ETAP

## 3.4 Arc Flash Reports

As for the short circuit component, there is no specific naming convention required for the arc flash reports. The only condition is that all the arc flash reports output by ETAP must be named starting with `AF_`. Some examples are as follows:

1. AF\_VCB\_PRES
2. AF\_VCBB\_ULT
3. AF\_VCBB\_Gen
  

Any report files contained in the ETAP project folder that do not start with the `AF_` will not be included in the short circuit table generation.

# 4\. Short Circuit

The short circuit tables provide a summary of fault levels at each bus under different system/switching configurations. The information used in the table is generated by ETAP in the form of short circuit report files. The Table Generator uses these files to extract specific fault level and impedance data for each bus and then organizes it into formatted tables. The fault values are presented in the phasor notation to make the tables compact.

The data is distributed into two sheets/tables namely, ‘Short Circuit Table’ and ‘Sequence Impedance Table’. The business logic of the program defines how the data is distributed between the three sheets. Refer to PSS Reference info for more details.

## 4.1 Prerequisites

For a successful run, the user must ensure the following:

1. If create reports is checked, the report files exist in the ETAP directory with the extension ‘SA2S’ and,
  1. Report files follow the naming convention defined in section 3.2.
  2. The report files are up to date.
2. If the create scenarios is checked,
  1. ETAP project and datahub are running.
  2. Study cases and switching configurations are set up in ETAP with the condition that the study cases are to be named as only the following:
    1. `SC`
    2. `SC_IEC`
      

## 4.2 Operation

The short circuit table generation is quite frankly, the simplest process to execute for the user. The user does not need to specify any options or exclusions since the short circuit table generation will consider and include all equipment that has been analyzed in the short circuit study.

The output produced is a single Excel file with two sheets, namely the ‘Short Circuit Table’ and the ‘Sequence Impedance Table’. The Short Circuit Table presents the fault level values in phasor notation at each equipment location whereas the Sequence Impedance Table outlines the impedance values at those locations in terms of resistance, reactance and magnitude.

The table may get quite large if the user has multiple scenarios. To allow for better handling, it is advised to divide a large table into separate tables grouped by similar scenarios.

# 5\. Device Duty

The device duty tables provide information regarding overdutied equipment. The information is generated from ETAP and subsequently used by the PSS Table Generator to organize it into tabular Excel Sheets to be added to the study reports later. The data is distributed into three different sheets namely, ANSI Momentary Table, ANSI Interrupting Table and IEC Interrupting Table. The business logic of the program defines how the data is distributed between the three sheets. Refer to PSS Reference info for more details.

## 5.1 Prerequisites

For a successful run, the user must ensure the following:

1. If create reports is checked, the report files exist in the ETAP directory with the extension ‘SA1S’ and/or ‘SI1S’ and,
  1. Report files follow the naming convention defined in section 3.3.
  2. The report files are up to date.
2. If the create scenarios is checked,
  1. ETAP project and datahub are running.
  2. Study cases and switching configurations are set up in ETAP with the condition that the study cases are to be named as only the following:
    1. `SC`
    2. `SC_IEC`
      

## 5.2 Study Options

For Device Duty, the interface of the PSS Table Generator offers four inputs, namely, `Include Switches`, `Use All Switching Configs`, `Add Series Ratings`, and `Mark Assumed Equipment`. The following sections briefly explain the purpose of each input.

### 5.2.1 Include Switches

When `Include Switches` is checked, the program will include switches/disconnects connected to the faulted buses in the device duty table. The switches will be appended at the end of the table following all the buses.

### 5.2.2 Use All Switching Configs

When `Use All Switching Configs` is checked, the program will use all switching configurations to create scenarios during the scenario creation process. If this option is unchecked, the program will only include the standard three switching configurations, namely, PRES, ULT and, GEN (if available).

### 5.2.3 Add Series Ratings

Enabling this option automatically highlights any panels and their branch breakers that are series rated in all tables. It also replaces the panel branch breaker rating with the corresponding series rating in the ANSI Interrupting Table. The series rating highlighting is shown in Figure 3 as orange.

In order to allow the program to find series rated panels and breakers, the user must ensure that the series rated equipment includes ‘SR’ (capitalization does not matter) in its comment field in ETAP. This includes the downstream series rated breakers of the panel.

Note: requires ETAP project and datahub running.

### 5.2.4 Mark Assumed Equipment

Enabling this option adds an asterisk (\*) to the IDs of equipment that have assumed ratings. This applies to all three tables. The asterisk can be seen for some equipment IDs in Figure 3.

In order to allow the program to find the equipment with assumed ratings, the user must ensure that the assumed rating equipment includes ‘assumed’ (capitalization does not matter) in its comment field in ETAP.

Note: requires ETAP project and datahub running.

Open 101946\_Device Duty Report1.png

![101946_Device Duty Report1.png](https://media-cdn.atlassian.com/file/050d8f5f-ff87-4ef9-9f28-43fd2fa9b5b0/image/cdn?allowAnimated=true&client=4ead6351-03ca-42fb-bbd7-8037c700c997&collection=contentId-3060957274&height=125&max-age=2592000&mode=full-fit&source=mediaCard&token=eyJhbGciOiJIUzI1NiJ9.eyJpc3MiOiI0ZWFkNjM1MS0wM2NhLTQyZmItYmJkNy04MDM3YzcwMGM5OTciLCJhY2Nlc3MiOnsidXJuOmZpbGVzdG9yZTpjb2xsZWN0aW9uOmNvbnRlbnRJZC0zMDYwOTU3Mjc0IjpbInJlYWQiXX0sImV4cCI6MTc3ODE4MTc0MiwibmJmIjoxNzc4MTc4ODYyLCJhYUlkIjoiNjEzNzlkMTUxNTZhYjMwMDcxNzVkZmNmIiwiaHR0cHM6Ly9pZC5hdGxhc3NpYW4uY29tL2FwcEFjY3JlZGl0ZWQiOmZhbHNlLCJhdXRoVHlwZSI6InNlc3Npb24ifQ.NKxaVnseEdbut9ORpLcNpszC1L2IB23pgyNioAB3B-g&width=692#media-blob-url=true&id=050d8f5f-ff87-4ef9-9f28-43fd2fa9b5b0&clientId=4ead6351-03ca-42fb-bbd7-8037c700c997&contextId=contentId-3060957274&collection=contentId-3060957274)

Figure 3. Table Generator Device Duty Report Output

## 5.3 Built-in Exclusions

The program will EXCLUDE any element types that are not listed in the table below.

| Momentary Table | Interrupting Table |
| :-- | :-- |
| SPSTDPSTPanelboardsSwitchboardsSwitchgear | All interrupting equipment is included. |

Once the user inputs all the required fields, the program can be executed by clicking the "Generate" button. The program shows a progress window indicating that the report generation process has started. The execution time typically ranges from a brief period of 3-6 seconds. During execution, any errors encountered will be displayed in a pop-up dialog. If the process completes without any errors, the application will automatically open the "Output Directory" where the generated file is located.

## 5.4 Switchgear Calculation

To calculate the switchgear symmetric rating, the program uses the following logic:

Open ID of SWGR starts with ‘MV’ or ‘LV’.png

![](https://media-cdn.atlassian.com/file/620ba1ad-4da8-4847-97ff-8c8aa1c135fb/image/cdn?allowAnimated=true&client=4ead6351-03ca-42fb-bbd7-8037c700c997&collection=contentId-3060957274&height=125&max-age=2592000&mode=full-fit&source=mediaCard&token=eyJhbGciOiJIUzI1NiJ9.eyJpc3MiOiI0ZWFkNjM1MS0wM2NhLTQyZmItYmJkNy04MDM3YzcwMGM5OTciLCJhY2Nlc3MiOnsidXJuOmZpbGVzdG9yZTpjb2xsZWN0aW9uOmNvbnRlbnRJZC0zMDYwOTU3Mjc0IjpbInJlYWQiXX0sImV4cCI6MTc3ODE4MTc0MiwibmJmIjoxNzc4MTc4ODYyLCJhYUlkIjoiNjEzNzlkMTUxNTZhYjMwMDcxNzVkZmNmIiwiaHR0cHM6Ly9pZC5hdGxhc3NpYW4uY29tL2FwcEFjY3JlZGl0ZWQiOmZhbHNlLCJhdXRoVHlwZSI6InNlc3Npb24ifQ.NKxaVnseEdbut9ORpLcNpszC1L2IB23pgyNioAB3B-g&width=500#media-blob-url=true&id=620ba1ad-4da8-4847-97ff-8c8aa1c135fb&clientId=4ead6351-03ca-42fb-bbd7-8037c700c997&contextId=contentId-3060957274&collection=contentId-3060957274)

Figure 5. Flowchart for the switchgear symmetric calculation

# 6\. Arc Flash

The arc flash report provides information regarding the incident energy and fault clearing time at various locations indicating the arc flash hazard. The information is generated from ETAP and subsequently used by the PSS Table Generator to organize it into a single tabular Excel sheet to be added to the study report later.

## 6.1 Prerequisites

For a successful run, the user must ensure the following:

1. If create reports is checked, the report files exist in the ETAP directory with the extension ‘AAFS’ and,
  1. Report files follow the naming convention defined in section 3.4.
  2. The report files are up to date.
2. If create scenarios is checked, the study cases, switching configurations (if any) and revisions (if any) are set up in ETAP with the condition that:
  1. The study cases are to be named as only the following:
    1. `AF_VCB`
    2. `AF_VCBB`
      

## 6.2 Study Options

For arc flash, the interface of the PSS Table Generator offers four inputs, namely, `Use SI Units`, `Incident Energy High / Critical`, and `Include Revisions`. The following sections briefly explain the purpose of each input.

### 6.2.1 Use SI Units

This option will convert all the ft-inch values to meter values. The multiplier used for conversion is:

**1 ft = 0.3048 m**

### 6.2.2 Incident Energy: High

The cells in the 'Incident Energy' column are highlighted in orange if their values exceed the threshold set in the `High` spinbox.

### 6.2.3 Incident Energy: Critical

The cells in the 'Incident Energy' column are highlighted in red if their values exceed the threshold set in the `Critical` spinbox.

### 6.2.4 Include Revisions

The revisions input is used only if the user opts to create scenarios. The radio button determines which revisions to include when generating arc flash scenarios in ETAP. If 'Base' is selected, only the base revision is used. If 'All' is selected, the program includes all revisions available in the ETAP project. When 'Only' is selected, the program uses only the revisions entered in the input box that appears when this option is chosen as can be seen in the image below. Multiple revisions can be listed in the text field with the `;` separator.

Open image-20241015-190835.png

![image-20241015-190835.png](https://media-cdn.atlassian.com/file/cfac2309-a558-4cb5-b76e-82cad8b7f42b/image/cdn?allowAnimated=true&client=4ead6351-03ca-42fb-bbd7-8037c700c997&collection=contentId-3060957274&height=125&max-age=2592000&mode=full-fit&source=mediaCard&token=eyJhbGciOiJIUzI1NiJ9.eyJpc3MiOiI0ZWFkNjM1MS0wM2NhLTQyZmItYmJkNy04MDM3YzcwMGM5OTciLCJhY2Nlc3MiOnsidXJuOmZpbGVzdG9yZTpjb2xsZWN0aW9uOmNvbnRlbnRJZC0zMDYwOTU3Mjc0IjpbInJlYWQiXX0sImV4cCI6MTc3ODE4MTc0MiwibmJmIjoxNzc4MTc4ODYyLCJhYUlkIjoiNjEzNzlkMTUxNTZhYjMwMDcxNzVkZmNmIiwiaHR0cHM6Ly9pZC5hdGxhc3NpYW4uY29tL2FwcEFjY3JlZGl0ZWQiOmZhbHNlLCJhdXRoVHlwZSI6InNlc3Npb24ifQ.NKxaVnseEdbut9ORpLcNpszC1L2IB23pgyNioAB3B-g&width=446#media-blob-url=true&id=cfac2309-a558-4cb5-b76e-82cad8b7f42b&clientId=4ead6351-03ca-42fb-bbd7-8037c700c997&contextId=contentId-3060957274&collection=contentId-3060957274)

Figure 6. Include revisions input

## 6.4 Output

The output of a successful run is an Excel table containing the incident energy information for various buses. The table is sorted in descending order of incident energy. The following image shows an example output.

Open ETAP Practice\_Arc Flash Report1.png

![ETAP Practice_Arc Flash Report1.png](https://media-cdn.atlassian.com/file/6e17018a-452b-4300-87eb-396aeb6f4c4e/image/cdn?allowAnimated=true&client=4ead6351-03ca-42fb-bbd7-8037c700c997&collection=contentId-3060957274&height=125&max-age=2592000&mode=full-fit&source=mediaCard&token=eyJhbGciOiJIUzI1NiJ9.eyJpc3MiOiI0ZWFkNjM1MS0wM2NhLTQyZmItYmJkNy04MDM3YzcwMGM5OTciLCJhY2Nlc3MiOnsidXJuOmZpbGVzdG9yZTpjb2xsZWN0aW9uOmNvbnRlbnRJZC0zMDYwOTU3Mjc0IjpbInJlYWQiXX0sImV4cCI6MTc3ODE4MTc0MiwibmJmIjoxNzc4MTc4ODYyLCJhYUlkIjoiNjEzNzlkMTUxNTZhYjMwMDcxNzVkZmNmIiwiaHR0cHM6Ly9pZC5hdGxhc3NpYW4uY29tL2FwcEFjY3JlZGl0ZWQiOmZhbHNlLCJhdXRoVHlwZSI6InNlc3Npb24ifQ.NKxaVnseEdbut9ORpLcNpszC1L2IB23pgyNioAB3B-g&width=760#media-blob-url=true&id=6e17018a-452b-4300-87eb-396aeb6f4c4e&clientId=4ead6351-03ca-42fb-bbd7-8037c700c997&contextId=contentId-3060957274&collection=contentId-3060957274)

Open ETAP Practice\_Arc Flash Report2.png

![ETAP Practice_Arc Flash Report2.png](https://media-cdn.atlassian.com/file/7c1caadc-f978-41f1-ba86-a561ccb31df7/image/cdn?allowAnimated=true&client=4ead6351-03ca-42fb-bbd7-8037c700c997&collection=contentId-3060957274&height=125&max-age=2592000&mode=full-fit&source=mediaCard&token=eyJhbGciOiJIUzI1NiJ9.eyJpc3MiOiI0ZWFkNjM1MS0wM2NhLTQyZmItYmJkNy04MDM3YzcwMGM5OTciLCJhY2Nlc3MiOnsidXJuOmZpbGVzdG9yZTpjb2xsZWN0aW9uOmNvbnRlbnRJZC0zMDYwOTU3Mjc0IjpbInJlYWQiXX0sImV4cCI6MTc3ODE4MTc0MiwibmJmIjoxNzc4MTc4ODYyLCJhYUlkIjoiNjEzNzlkMTUxNTZhYjMwMDcxNzVkZmNmIiwiaHR0cHM6Ly9pZC5hdGxhc3NpYW4uY29tL2FwcEFjY3JlZGl0ZWQiOmZhbHNlLCJhdXRoVHlwZSI6InNlc3Npb24ifQ.NKxaVnseEdbut9ORpLcNpszC1L2IB23pgyNioAB3B-g&width=760#media-blob-url=true&id=7c1caadc-f978-41f1-ba86-a561ccb31df7&clientId=4ead6351-03ca-42fb-bbd7-8037c700c997&contextId=contentId-3060957274&collection=contentId-3060957274)

Figure 7: Table Generator Arc Flash Report Output

# 7\. Exclusions

PSS Table Generator provides a feature that allows users to filter elements by their IDs. Users can exclude specific elements from the tables by entering relevant keywords in the "Exclusions" section of the application. Two input fields are available: `Element ID Starting With` and `Element ID Containing`. The following sections provide a brief overview of each input option.

## 7.1 Element ID Starting With

This option allows users to filter elements based on their IDs. To use this feature, enter words separated by colons in the input field. The program will exclude any elements whose IDs begin with any of the specified words.

Note: capitalization is case-sensitive.

## 7.2 Element ID Containing

This option allows the user to elements based on their IDs. The input requires words separated by a colon. The program will exclude any elements whose ID contains any of the words in the input field. The capitalization matters.

Note: capitalization is case-sensitive.

# 8\. Menu Bar

# 9\. Uninstallation

To uninstall the application, the user can go to the installation directory and run the uninstaller. Alternatively, the user can uninstall from the programs option in the control panel.

# 10\. References

1. [PSS Table Generator - GitHub Source Code](https://github.com/primeeng-adil/pss-table-generator "https://github.com/primeeng-adil/pss-table-generator")
2. [PSS Procedures, Guidelines and General Notes - Power Systems Studies - Confluence (atlassian.net)](https://primeeng.atlassian.net/wiki/spaces/PST/pages/58621985 "https://primeeng.atlassian.net/wiki/spaces/PST/pages/58621985")