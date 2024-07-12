import paramiko
import getpass
from office365.runtime.auth.user_credential import UserCredential
from office365.sharepoint.client_context import ClientContext
from office365.sharepoint.files.file import File

def search_interface_description(ip_list, username, password, port, description):
    try:
        for ip in ip_list:
            client = paramiko.SSHClient()
            client.set_missing_host_key_policy(paramiko.AutoAddPolicy())
            client.connect(ip, username=username, password=password, port=port, timeout=5)

            command = f"sh int des | include {description}"
            stdin, stdout, stderr = client.exec_command(command)

            output = stdout.read().decode('utf-8')
            error = stderr.read().decode('utf-8')
            
            if error:
                print(f"Error connecting to {ip}: {error}")
            else:
                print(f"Results for {ip} for search query '{description}':")
                print(output)
            
            client.close()
    except Exception as e:
        print(f"Error connecting or sending command: {str(e)}")

def search_by_mac(ip_list, username, password, port, mac_address):
    try:
        for ip in ip_list:
            client = paramiko.SSHClient()
            client.set_missing_host_key_policy(paramiko.AutoAddPolicy())
            client.connect(ip, username=username, password=password, port=port, timeout=5)

            command = f"sh mac add add {mac_address}"
            stdin, stdout, stderr = client.exec_command(command)

            output = stdout.read().decode('utf-8')
            error = stderr.read().decode('utf-8')
            
            if error:
                print(f"Error connecting to {ip}: {error}")
            else:
                print(f"Results for {ip} for search query with MAC address '{mac_address}':")
                print(output)
            
            client.close()
    except Exception as e:
        print(f"Error connecting or sending command: {str(e)}")

def backup_config(ip_list, username, password, port, sharepoint_site_url, sharepoint_folder, sharepoint_username, sharepoint_password):
    try:
        for ip in ip_list:
            client = paramiko.SSHClient()
            client.set_missing_host_key_policy(paramiko.AutoAddPolicy())
            client.connect(ip, username=username, password=password, port=port, timeout=5)

            command = "show running-config"
            stdin, stdout, stderr = client.exec_command(command)

            output = stdout.read().decode('utf-8')
            error = stderr.read().decode('utf-8')
            
            if error:
                print(f"Error connecting to {ip}: {error}")
            else:
                backup_filename = f"{ip}_backup.txt"
                with open(backup_filename, "w") as file:
                    file.write(output)
                print(f"Configuration backup for {ip} saved to {backup_filename}")

                try:
                    ctx = ClientContext(sharepoint_site_url).with_credentials(UserCredential(sharepoint_username, sharepoint_password))
                    target_folder = ctx.web.get_folder_by_server_relative_url(sharepoint_folder)
                    with open(backup_filename, 'rb') as content_file:
                        file_content = content_file.read()
                        target_folder.upload_file(backup_filename, file_content).execute_query()
                    print(f"Configuration backup for {ip} uploaded to SharePoint at {sharepoint_folder}/{backup_filename}")
                except Exception as sharepoint_error:
                    print(f"SharePoint error: {sharepoint_error}")
            
            client.close()
    except Exception as e:
        print(f"Error connecting or sending command: {str(e)}")

def backup_firewall_config(ip_list, username, password, port, sharepoint_site_url, sharepoint_folder, sharepoint_username, sharepoint_password):
    try:
        for ip in ip_list:
            client = paramiko.SSHClient()
            client.set_missing_host_key_policy(paramiko.AutoAddPolicy())
            client.connect(ip, username=username, password=password, port=port, timeout=5)

            command = "show full-configuration"
            stdin, stdout, stderr = client.exec_command(command)

            output = stdout.read().decode('utf-8')
            error = stderr.read().decode('utf-8')
            
            if error:
                print(f"Error connecting to {ip}: {error}")
            else:
                backup_filename = f"{ip}_firewall_backup.txt"
                with open(backup_filename, "w") as file:
                    file.write(output)
                print(f"Firewall configuration backup for {ip} saved to {backup_filename}")

                try:
                    ctx = ClientContext(sharepoint_site_url).with_credentials(UserCredential(sharepoint_username, sharepoint_password))
                    target_folder = ctx.web.get_folder_by_server_relative_url(sharepoint_folder)
                    with open(backup_filename, 'rb') as content_file:
                        file_content = content_file.read()
                        target_folder.upload_file(backup_filename, file_content).execute_query()
                    print(f"Firewall configuration backup for {ip} uploaded to SharePoint at {sharepoint_folder}/{backup_filename}")
                except Exception as sharepoint_error:
                    print(f"SharePoint error: {sharepoint_error}")
            
            client.close()
    except Exception as e:
        print(f"Error connecting or sending command: {str(e)}")

def main():
    while True:
        print("Menu:")
        print("1. Search for Tag description on Cisco Switch device")
        print("2. Search by MAC address")
        print("3. Backup Cisco Switch configuration")
        print("4. Backup Fortinet Firewall configuration")
        print("5. Return to main menu")

        choice = input("Please choose a menu by typing the number: ")

        if choice == '1':
            ip_addresses = []
            while True:
                ip = input("Please enter the IP Address of the Cisco Switch device (enter 0 to stop): ")
                if ip == '0':
                    break
                ip_addresses.append(ip)

            if not ip_addresses:
                print("No IP Address entered")
                continue

            username = input("Please enter the SSH username: ")
            password = getpass.getpass("Please enter the SSH password: ")
            port = 22
            description = input("Please enter the description to search for: ")

            search_interface_description(ip_addresses, username, password, port, description)

        elif choice == '2':
            ip_addresses = []
            while True:
                ip = input("Please enter the IP Address of the Cisco Switch device (enter 0 to stop): ")
                if ip == '0':
                    break
                ip_addresses.append(ip)

            if not ip_addresses:
                print("No IP Address entered")
                continue

            username = input("Please enter the SSH username: ")
            password = getpass.getpass("Please enter the SSH password: ")
            port = 22
            mac_address = input("Please enter the MAC Address to search for: ")

            search_by_mac(ip_addresses, username, password, port, mac_address)

        elif choice == '3':
            ip_addresses = []
            while True:
                ip = input("Please enter the IP Address of the Cisco Switch device (enter 0 to stop): ")
                if ip == '0':
                    break
                ip_addresses.append(ip)

            if not ip_addresses:
                print("No IP Address entered")
                continue

            username = input("Please enter the SSH username: ")
            password = getpass.getpass("Please enter the SSH password: ")
            port = 22
            sharepoint_site_url = input("Please enter the SharePoint site URL: ")
            sharepoint_folder = input("Please enter the SharePoint folder path: ")
            sharepoint_username = input("Please enter the SharePoint username: ")
            sharepoint_password = getpass.getpass("Please enter the SharePoint password: ")

            backup_config(ip_addresses, username, password, port, sharepoint_site_url, sharepoint_folder, sharepoint_username, sharepoint_password)

        elif choice == '4':
            ip_addresses = []
            while True:
                ip = input("Please enter the IP Address of the Fortinet Firewall device (enter 0 to stop): ")
                if ip == '0':
                    break
                ip_addresses.append(ip)

            if not ip_addresses:
                print("No IP Address entered")
                continue

            username = input("Please enter the SSH username: ")
            password = getpass.getpass("Please enter the SSH password: ")
            port = 22
            sharepoint_site_url = input("Please enter the SharePoint site URL: ")
            sharepoint_folder = input("Please enter the SharePoint folder path: ")
            sharepoint_username = input("Please enter the SharePoint username: ")
            sharepoint_password = getpass.getpass("Please enter the SharePoint password: ")

            backup_firewall_config(ip_addresses, username, password, port, sharepoint_site_url, sharepoint_folder, sharepoint_username, sharepoint_password)

        elif choice == '5':
            print("Returning to main menu")
            break

        else:
            print("Invalid menu choice")
            continue

if __name__ == "__main__":
    main()
