---

# Tableau Server Role Synchronization Script

This Python script synchronizes user roles between a Source Tableau Server and a Target Tableau Server. It updates the user roles on the Target server to match those on the Source server, logs the changes, and sends notification emails to the affected users and administrators.

## Features

- **Role Synchronization**: The script synchronizes user roles from the Source Tableau Server to the Target Tableau Server.
- **Logging**: The script logs the sites and users processed during the synchronization.
- **Email Notifications**: 
  - Sends emails to users whose roles were updated on the Target server, detailing the changes.
  - Sends a summary email to the admin team with the details of all updated users.

## Prerequisites

- **Python Libraries**: 
  - `tableauserverclient`
  - `pandas`
  - `smtplib`
  - `email`
  - `pytz`
  - `logging`
  
  Install these libraries using `pip`:
  ```bash
  pip install tableauserverclient pandas pytz
  ```

- **SMTP Server**: Ensure that you have access to an SMTP server to send emails.

## Configuration

Before running the script, update the following variables with your specific details:

- **Tableau Server URLs and Site IDs**:
  - `source_server_url`: URL of the Source Tableau Server.
  - `source_site_id`: Site ID of the Source server (leave empty if not using multi-site).
  - `target_server_url`: URL of the Target Tableau Server.
  - `target_site_id`: Site ID of the Target server (leave empty if not using multi-site).

- **Authentication Details**:
  - `source_username`: Username for the Source server.
  - `source_password`: Password for the Source server.
  - `target_username`: Username for the Target server.
  - `target_password`: Password for the Target server.

- **SMTP Configuration**:
  - `smtp_host`: Hostname of the SMTP server.
  - `smtp_port`: Port of the SMTP server.
  - `From`: Email address to use as the sender.
  - `To`: Recipient's email address.
  - `Bcc`: BCC recipient email address (if needed).
  - `admin_dl`: Email address of the admin distribution list.

## Usage

1. **Run the Script**: Execute the script in your Python environment:
    ```bash
    python main.py
    ```

2. **Process Flow**:
   - The script logs into both the Source and Target Tableau Servers.
   - It retrieves the list of users from each site on both servers.
   - User roles on the Target server are updated to match those on the Source server.
   - A detailed log of changes is generated and saved to an Excel file (`updated_user_roles.xlsx`).
   - Emails are sent to the users and admin team with the updated role information.

## Output

- **Excel File**: `updated_user_roles.xlsx` contains the details of all users whose roles were updated, including:
  - User Name
  - Site Name
  - Previous Role
  - New Role
  - Email
  - Full Name
  - Last Login Time

- **Email Notifications**: 
  - Users receive an email with details about their updated roles.
  - The admin team receives a summary email listing all the changes made.

## Logging

- The script logs activities and errors using the Python `logging` module. Ensure that logging is appropriately configured to capture important information.

## Troubleshooting

- **Connection Issues**: Verify the URLs and credentials for the Tableau servers.
- **Email Issues**: Ensure the SMTP server is correctly configured and reachable from the environment where the script is running.

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

---
