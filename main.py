import tableauserverclient as TSC
import pandas as pd
import smtplib
from email.message import EmailMessage
import pytz
import logging

# Define the server URLs and site IDs for both Prod and QA
prod_server_url = "https://prd.abc.com/"
prod_site_id = ""
qa_server_url = "https://qa.abc.com/"
qa_site_id = ""

# Define the authentication tokens for both servers
prod_username = ""
prod_password = ""

qa_username = ""
qa_password = ""

updated_users_global = []
updated_users_excel = []

smtp_host = 'mail.abc.com'
smtp_port = 25
From = 'admin@abc.com'
To = ''
Bcc = ''
admin_dl = 'admin@abc.com'


def main():
    global updated_users_excel
    # Connect to both servers using the Tableau Server Client library
    prod_auth = TSC.TableauAuth(prod_username, prod_password)
    prod_server = TSC.Server(prod_server_url)
    prod_server.add_http_options({'verify': False})
    prod_server.use_server_version()

    qa_auth = TSC.TableauAuth(qa_username, qa_password)
    qa_server = TSC.Server(qa_server_url)
    qa_server.add_http_options({'verify': False})
    qa_server.use_server_version()

    updated_users_local = []

    with prod_server.auth.sign_in(prod_auth):
        with qa_server.auth.sign_in(qa_auth):

            # Retrieve all sites from both Prod and QA
            prod_sites, prod_pagination = prod_server.sites.get()
            print("Sites from Prod:", prod_sites)
            qa_sites, qa_pagination = qa_server.sites.get()
            print("Sites from QA:", qa_sites)

            # Loop through each site on the Prod server
            for prod_site in prod_sites:
                logging.info(f"Checking for Prod Site: {prod_site.name}")
                prod_server.auth.switch_site(prod_site)
                prod_all_users = [users for users in TSC.Pager(prod_server.users)]

                # Create a dictionary of the Prod users with their site roles
                prod_user_roles = {user.name: user.site_role for user in prod_all_users if user.site_role
                                   not in ['ServerAdministrator', 'SiteAdministratorCreator',
                                           'SiteAdministratorExplorer']}

                # Retrieve the corresponding site on the QA server
                qa_site = next((site for site in qa_sites if site.name == prod_site.name), None)
                if not qa_site:
                    continue
                qa_server.auth.switch_site(qa_site)
                logging.info(f"Checking for QA Site: {qa_site.name}")
                qa_all_users = [users for users in TSC.Pager(qa_server.users)]

                # Loop through the QA users and update their site roles if necessary
                for user in qa_all_users:
                    print("User from QA:", user.name)
                    user_name = user.name
                    user_role = user.site_role

                    # Check if the user has a different role in Prod
                    if user_name in prod_user_roles and prod_user_roles[user_name] != user_role:
                        # Update the user's role in QA to match their role in Prod
                        prev_role = user_role
                        user.site_role = prod_user_roles[user_name]
                        qa_server.users.update(user)
                        if user.last_login is not None and user.last_login.tzinfo is not None:
                            # Convert to CST timezone
                            cst_tz = pytz.timezone('US/Central')
                            dt_cst = user.last_login.astimezone(cst_tz).replace(tzinfo=None)
                        else:
                            dt_cst = "Never Logged In"
                        updated_users_local.append(
                            {'USER_NAME': user_name, 'SITE_NAME': qa_site.name, 'PREV_ROLE': prev_role,
                             'NEW_ROLE': user.site_role, 'LAST_LOGIN': dt_cst, 'EMAIL': user.email})

                        updated_users_excel.append(
                            {'USER_NAME': user_name, 'SITE_NAME': qa_site.name, 'PREV_ROLE': prev_role,
                             'NEW_ROLE': user.site_role, 'EMAIL': user.email,
                             'FULL NAME': user.fullname,
                             'LAST_LOGIN': dt_cst})
            qa_server.auth.sign_out()
        prod_server.auth.sign_out()
    # Write the updated user details to an Excel file using Pandas
    df = pd.DataFrame(updated_users_excel, columns=['USER_NAME', 'SITE_NAME', 'PREV_ROLE', 'NEW_ROLE', 'EMAIL',
                                                    'FULL NAME', 'LAST_LOGIN'])
    df.to_excel('updated_user_roles.xlsx', index=False)
    return updated_users_local


def email_users(updated_users_local):
    # Convert the list of updated user details to a Pandas DataFrame
    users_email_df = pd.DataFrame(updated_users_local)

    # Group the users by email
    grouped_users = users_email_df.groupby('EMAIL')

    # Loop through each group and send email to the corresponding user
    try:
        for email, group in grouped_users:
            # Set up email message
            msg = EmailMessage()
            msg['Subject'] = 'Notification: QA Tableau Server Access Update'
            msg['From'] = From
            msg['To'] = email  # replace To with email to send emails to users.
            msg['Bcc'] = ''
            user_name = group['USER_NAME'].unique()[0]
            # Create HTML table with updated role information for the user
            email_body = group.to_html(na_rep="", index=False, render_links=True,
                                       escape=False).replace('<th>', '<th style="color:white; '
                                                                     'background-color:#180e62">')
            msg.set_content(email_body)
            msg.add_alternative(f"""
                        <!DOCTYPE html>
                        <html>
                          <body style='font-family: Merriweather; font-size: 11px;'>
                            <p style='color: #44546A;'>Dear <b>{user_name}</b>, </p>
                            <p style='color: #44546A;'>This is to inform you that your Tableau Server ({qa_server_url}) 
                            access has been updated to match with Production Server.</p>
                            <p style='color: #44546A;'>
                              <b>Details of your updated access are as follows:</b>
                            </p>{email_body} <br>
                            <p style='color: #44546A;'>Thank you for your understanding and cooperation in this matter.</p>
                            <p style='color: #44546A;'>Regards, <br>QA Access Syncing Automation</p>
                          </body>
                        </html>""", subtype='html')
            with smtplib.SMTP(smtp_host, smtp_port) as smtp:
                smtp.send_message(msg)
    except Exception as e:
        logging.info(f"Error sending email: {str(e)}")


def email_admin(updated_users_excel):
    # Convert the list of updated user details to a Pandas DataFrame
    users_email_df = pd.DataFrame(updated_users_excel)
    try:
        # Set up email message
        msg = EmailMessage()
        msg['Subject'] = 'Notification: QA Tableau Server Access Syncing Update'
        msg['From'] = From
        msg['To'] = admin_dl  # replace To with email to send emails to users.
        msg['Bcc'] = ''
        # Create HTML table with updated role information for the user
        email_body = users_email_df.to_html(na_rep="", index=False, render_links=True,
                                           escape=False).replace('<th>', '<th style="color:white; '
                                                                         'background-color:#180e62">')
        msg.set_content(email_body)
        msg.add_alternative(f""" <!DOCTYPE html> <html> <body style='font-family: Merriweather; font-size: 11px;'> <p 
        style='color: #44546A;'>Hi Admin Team, </p> <p style='color: #44546A;'>This is to inform you that accesses of 
        below users were synced on QA Tableau Server ({qa_server_url}) with Prod Server.</p> 
                                <p style='color: #44546A;'>
                                  <b>Details of updated accesses are as follows:</b>
                                </p>{email_body} <br>
                                <p style='color: #44546A;'>Regards, <br>Access Syncing Automation</p>
                              </body>
                            </html>""", subtype='html')
        with smtplib.SMTP(smtp_host, smtp_port) as smtp:
            smtp.send_message(msg)
    except Exception as e:
        logging.info(f"Error sending email: {str(e)}")


if __name__ == '__main__':
    updated_users_global = main()
    email_users(updated_users_global)
    email_admin(updated_users_excel)
