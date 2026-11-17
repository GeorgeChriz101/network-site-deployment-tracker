# Network Site Deployment Schedule and Resource Tracker

A PMO portfolio project simulating the planning and coordination of five cellular network deployment sites.

## Project Overview

This project demonstrates how a project coordination or PMO team can manage a multi-site network deployment schedule.

The tracker covers five sample cellular network sites and five major deployment stages:

- Site Survey
- Permitting
- Equipment Delivery
- Installation
- Acceptance

This creates a total of 25 tracked milestones.

The project also includes six simulated delay and resource-conflict scenarios to demonstrate schedule impact analysis, corrective action tracking, resequencing, and change control.

## Tools Used

- Python
- pandas
- Excel
- openpyxl
- Microsoft Word
- Git

## Key Features

- 5 sample network deployment sites
- 25 scheduled milestones
- Milestone owners
- Planned start and end dates
- Milestone dependencies
- Status and percent-complete tracking
- 6 simulated delay/resource-conflict scenarios
- Proposed schedule recovery actions
- PMO action log
- Example change request
- Excel project workbook
- Supporting PMO documentation

## Project Structure

network-site-deployment-tracker/

- data/
  - deployment_schedule.csv
  - delay_scenarios.csv
  - action_log.csv
  - change_request.csv
- docs/
  - action_log_guide.md
  - change_control_process.md
  - data_dictionary.md
  - project_methodology.md
  - project_summary.md
  - risk_and_issue_guide.md
  - schedule_management_guide.md
- reports/
  - network_site_deployment_tracker.xlsx
- scripts/
  - build_project.py
- README.md
- requirements.txt
- .gitignore

## Project Workflow

Site Planning  
→ Milestone Scheduling  
→ Assign Owners  
→ Track Dependencies  
→ Monitor Progress  
→ Identify Delays and Resource Conflicts  
→ Create Corrective Actions  
→ Evaluate Schedule Changes  
→ Update Stakeholders

## Delay and Resource Scenarios

The project tests situations such as:

- Permit approval delays
- Equipment delivery delays
- Installation crew conflicts
- Engineer scheduling conflicts
- Acceptance resource availability
- Combined schedule and resource conflicts

## Action Management

The action log tracks:

- Action ID
- Site
- Action
- Owner
- Due date
- Status

## Change Control

The project includes an example change request documenting:

- Requested change
- Reason
- Schedule impact
- Cost impact
- Decision status

## Excel Workbook

The Excel workbook contains:

- Deployment Schedule
- Delay Scenarios
- Action Log
- Change Request

## Run the Project

Activate the environment:

    source .venv/bin/activate

Run:

    python scripts/build_project.py

## Purpose

The project demonstrates practical PMO and project coordination skills in a network deployment environment, including scheduling, milestone tracking, dependency management, resource coordination, issue tracking, corrective actions, and change control.

All project information is synthetic and created for portfolio demonstration purposes.
