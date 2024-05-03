# -*- coding: utf-8 -*-
"""
Created on Fri Feb  9 14:11:16 2024

@author: EXU552
"""
import requests
import mysql.connector
from mysql.connector import Error
import configparser
import time

MAX_RETRIES = 3

with open("logs_insert.log", "w"):
    pass

token = None
connection = None

data = {
    'grant_type': 'client_credentials',
    'client_id': 'uri:a947p01.chargeback.infrabel.be-ucmdb.api',
    'client_secret': 'pTyQ1yf_zBQnDahnPdNBlvW7mEyP__wKxXpiXfgk',
    'scope': 'https://ucmdb-api.infrabel.be'
}

def get_token(data=data):
    response = requests.post('https://claim.infrabel.be/adfs/oauth2/token', data=data)
    response_json = response.json()
    return response_json.get("access_token", None)

def api_request(url, params, headers):
    global token
    if not token:
        token = get_token()
    headers['authorization'] = "Bearer {}".format(token)
    response = requests.get(url, params=params, headers=headers)
    if response.status_code in [401, 500]:
        token = get_token()
        headers['authorization'] = "Bearer {}".format(token)
        response = requests.get(url, params=params, headers=headers)
    return response.json()

def reset_db(connection):
    tables_to_reset = ['A1788_application', 'A1788_environment', 'A1788_software_server', 'A1788_computer_system', 'A1788_ip', 'A1788_software_server_environment', 'A1788_computer_system_environment']
    for table_to_reset in tables_to_reset:
        try:
            cursor = connection.cursor()
            delete_query = f"DELETE FROM {table_to_reset}"
            cursor.execute(delete_query)
            connection.commit()
            print(f"Table {table_to_reset} emptied successfully.")
            alter_query = f"ALTER TABLE {table_to_reset} AUTO_INCREMENT = 1"
            cursor.execute(alter_query)
            connection.commit()
        except Error as e:
            print(f"Error emptying table {table_to_reset}: {e}")
        finally:
            cursor.close()

def establish_db_connection():
    global connection
    try:
        config = configparser.ConfigParser()
 
        # Reading the configuration file
        config.read('config.ini')
     
        # Accessing values from the configuration file
        conf_host = config.get('Database', 'host')
        conf_database = config.get('Database', 'database')
        conf_user = config.get('Database', 'user')
        conf_password = config.get('Database', 'password')
        conf_port = config.get('Database', 'port')
        conf_auth_plugin = config.get('Database', 'auth_plugin')
        
        connection = mysql.connector.connect(
            host = conf_host,
            database = conf_database,
            user = conf_user,
            password = conf_password,
            port = conf_port,
            auth_plugin = conf_auth_plugin
        )
        if connection.is_connected():
            print("Database connection established successfully.")
        else:
            print("Failed to establish database connection.")
    except Error as e:
        print(f"Error establishing database connection: {e}")

def close_db_connection():
    global connection
    if connection is not None and connection.is_connected():
        connection.close()
        print("Database connection closed.")

def get_app_codes():
    params = {"RequestFor": "Application", "Detail": "B", "Identifiers": "All"}
    headers = {"content-type": "application/json"}
    response_data = api_request("https://ucmdb.infrabel.be/api/dbdata/procedures/pGetCiInfoVdelta", params, headers)
    all_codes = []
    for app in response_data:
        code = app.get("code", "")
        if code != "UNKNOWN":
            if app.get("status", "") == "Operational":
                all_codes.append(code)  
    return all_codes

def get_app(code):
    params = {"RequestFor": "Application", "Detail": "B", "Identifiers": code}
    headers = {"content-type": "application/json"}
    response_data = api_request("https://ucmdb.infrabel.be/api/dbdata/procedures/pGetCiInfoVdelta", params, headers)
    return [[app.get("identification", ""), app.get("nameFr", ""), app.get("nameNl", ""), app.get("nameEn", ""), app.get("status", "")] for app in response_data]

def get_env(code):
    params = {"RequestFor": "Application", "Detail": "D", "Identifiers": code, "Environments": "All"}
    headers = {"content-type": "application/json"}
    response_data = api_request("https://ucmdb.infrabel.be/api/dbdata/procedures/pGetCiInfoVdelta", params, headers)
    env_data = []
    for app in response_data:
        environments = app.get("environments", [])
        if environments is not None:
            for env in environments:
                if env.get("status", "") != "Destroyed":
                    env_data.extend([env.get("environmentId", ""), env.get("identification", ""), env.get("environmentType", ""), env.get("status", "")])
    return env_data

def get_cs(code):
    params = {"RequestFor": "Application", "Detail": "D", "Identifiers": code, "ComputerSystems": "All"}
    headers = {"content-type": "application/json"}
    response_data = api_request("https://ucmdb.infrabel.be/api/dbdata/procedures/pGetCiInfoVdelta", params, headers)
    cs_data = []
    for app in response_data:
        environments = app.get("environments", [])
        if environments is not None:
            for env in environments:
                env_guid = env.get('environmentId', '')
                computer_systems = env.get("computerSystems", [])
                if computer_systems is not None:
                    cs_data.extend([[cs.get("computerSystemId", ""), cs.get("identification", ""), cs.get("status", ""), env_guid] for cs in computer_systems])
    return cs_data

def get_ss(code):
    params = {'RequestFor': 'Application', 'Detail': 'D', 'Identifiers': code, 'SoftwareServers': 'All'}
    headers = {'content-type': "application/json"}
    response_data_1 = api_request('https://ucmdb.infrabel.be/api/dbdata/procedures/pGetCiInfoVdelta', params, headers)
    ss_data = []
    ss_identifiers = []
    for app in response_data_1:
        environments = app.get('environments', [])
        if environments is not None:
            for env in environments:
                software_servers = env.get('softwareServers', [])
                if software_servers is not None:
                    for ss in software_servers:
                        identifier = ss.get('identification', '')
                        if len(identifier) >= 3 and '%' not in identifier:
                            if identifier not in ss_identifiers:
                                ss_identifiers.append(identifier)
    
    all_ss = []
    for identifier in ss_identifiers:
        params = {'RequestFor': 'SoftwareServer', 'Detail': 'D', 'Identifiers': identifier, 'Environments': 'All', 'ComputerSystems': 'All'}
        headers = {'content-type': "application/json"}
        response_data_2 = api_request('https://ucmdb.infrabel.be/api/dbdata/procedures/pGetCiInfoVdelta', params, headers)
        for ss in response_data_2:
            if ss not in all_ss:
                all_ss.append(ss)
    for ss in all_ss:
        ss_id = ss.get('softwareServerId', '')
        ss_identification = ss.get('identification', '')
        ss_short_desc = ss.get('shortDescription', '')
        ss_status = ss.get('status', '')
        environments = ss.get('environments', [])
        computer_systems = ss.get('computerSystems', [])
        if environments is not None and computer_systems is not None:
            for env in environments:
                env_name = env.get('name', '')
                env_guid = env.get('environmentId', '')
                if code in env_name:
                    for cs in computer_systems:
                        cs_id = cs.get('computerSystemId', '')
                        entry = [ss_id, ss_identification, ss_short_desc, ss_status, env_guid, cs_id]
                        if entry not in ss_data:
                            ss_data.append(entry)      
    return ss_data 

def get_ip(code):
    params = {"RequestFor": "Application", "Detail": "D", "Identifiers": code, "ComputerSystems": "All"}
    headers = {"content-type": "application/json"}
    response_data = api_request("https://ucmdb.infrabel.be/api/dbdata/procedures/pGetCiInfoVdelta", params, headers)
    ip_data = []
    for app in response_data:
        environments = app.get("environments", [])
        if environments is not None:
            for env in environments:
                computer_systems = env.get("computerSystems", [])
                if computer_systems is not None:
                    ip_data.extend([[cs.get("ipAddress", ""), cs.get("computerSystemId", "")] for cs in computer_systems])
    return ip_data

def get_app_id(code, connection):
    try:
        cursor = connection.cursor()
        cursor.execute(f"SELECT app_id FROM A1788_application WHERE a_code LIKE '{code}'")
        result = cursor.fetchone()
        if result:
            return result[0]
        else:
            return None
    except Error as e:
        print(f"Error getting application ID: {e}")
        return None
            
def get_softwareserver_environment_data(connection):
    try:
        cursor = connection.cursor()
        
        # Fetching environment and software server IDs based on the provided code
        select_query = "SELECT A1788_environment.env_id, A1788_software_server.ss_id FROM A1788_environment INNER JOIN A1788_software_server ON A1788_environment.guid = A1788_software_server.env_guid"
        cursor.execute(select_query)
        ss_env_data = cursor.fetchall()
        
        return ss_env_data

    except Error as e:
        print(f"Error getting software server environment data: {e}")
        return None
            
def get_computersystem_environment_data(connection):
    try:
        cursor = connection.cursor()     
        # Fetching environment and software server IDs based on the provided code
        select_query = "SELECT A1788_environment.env_id, A1788_computer_system.cs_id FROM A1788_environment INNER JOIN A1788_computer_system ON A1788_environment.guid = A1788_computer_system.env_guid"
        cursor.execute(select_query)
        cs_env_data = cursor.fetchall() 
        
        return cs_env_data

    except Error as e:
        print(f"Error getting computer system environment data: {e}")
        return None
                        
def insert_data_into_db(data, table_name, code):
    table_columns = {
        'A1788_application': ('a_code', 'name_fr', 'name_nl', 'name_en', '_status'),
        'A1788_environment': ('guid', '_name', '_type', '_status', 'app_id'),
        'A1788_computer_system': ('guid', '_name', '_status', 'env_guid'),
        'A1788_software_server': ('guid', 'identification', 'short_desc', '_status', 'env_guid', 'cs_id'),
        'A1788_ip': ('ip_address', 'cs_id'),
        'A1788_software_server_environment': ('env_id', 'ss_id'),
        'A1788_computer_system_environment': ('env_id', 'cs_id')
    }
    
    def insert_into_table(table_name, data, columns, code):
        global connection
        cursor = connection.cursor()
        try:
            if table_name == 'A1788_environment':
                # Getting applicationId
                app_id = get_app_id(code, connection)
                if app_id is None:
                    print(f"Application ID not found for code {code}")
                    return
                # Modifying data to include applicationId
                data = [(row[0], row[1], row[2], row[3], app_id) for row in data]
                
            elif table_name == 'A1788_software_server':
                cursor.execute("SELECT cs_id, guid FROM A1788_computer_system")
                cs_id_mapping = {row[1]: row[0] for row in cursor.fetchall()}
                data = [(row[0], row[1], row[2], row[3], row[4], cs_id_mapping.get(row[5])) for row in data]
                           
            elif table_name == 'A1788_ip':
                cursor.execute("SELECT cs_id, guid FROM A1788_computer_system")
                cs_id_mapping = {row[1]: row[0] for row in cursor.fetchall()}
                new_data = []
                for row in data:
                    ip_addresses = []
                    if row[0] is not None:
                        ip_addresses = row[0].split(';')
                    for ip_address in ip_addresses:
                        new_data.append((ip_address, cs_id_mapping.get(row[1])))
                data = new_data
            
            if data:
                insert_query = f"INSERT INTO {table_name} ({', '.join(columns)}) VALUES ({', '.join(['%s'] * len(columns))})"
                cursor.executemany(insert_query, data)
                connection.commit()
                print(f"{code}: Insert in {table_name} table successful.")
            else:
                print(f"No new data to insert in {table_name}.")
            
        except Error as e:
            with open("logs_insert.log", "a") as logfile:
               logfile.write(f"{code}: Insertion into {table_name} failed: {e}\n")
            print(f"{code}: Insertion into {table_name} failed: {e}")
        finally:
            cursor.close()

    if connection is not None:
        try:
            if table_name in table_columns:
                insert_into_table(table_name, data, table_columns[table_name], code)
            else:
                print(f"{code}: Table '{table_name}' not found.")
        except Error as e:
            print(f"{code}: Insertion into {table_name} failed: {e}")
    else:
        print("Database connection not established.")

if __name__ == "__main__":
    retry_count = 0
    while retry_count < MAX_RETRIES:
        establish_db_connection() 
        reset_db(connection)
        print("---------------------------------------------------------------")
        if connection is not None:
            try:
                all_codes = get_app_codes()
                all_codes.remove("A2024")
                remaining = len(all_codes) - 1
                for code in all_codes:
                    sse_before_insert = get_softwareserver_environment_data(connection)
                    cse_before_insert = get_computersystem_environment_data(connection)
                    
                    app_data = get_app(code)
                    insert_data_into_db(app_data, 'A1788_application', code)
                    
                    env_data = get_env(code)
                    insert_data_into_db(env_data, 'A1788_environment', code)
                    
                    cs_data = get_cs(code)
                    insert_data_into_db(cs_data, 'A1788_computer_system', code)
                    
                    ss_data = get_ss(code)
                    insert_data_into_db(ss_data, 'A1788_software_server', code)
                                    
                    ip_data = get_ip(code)
                    insert_data_into_db(ip_data, 'A1788_ip', code)
                    
                    sse_after_insert = get_softwareserver_environment_data(connection)
                    cse_after_insert = get_computersystem_environment_data(connection)
               
                    sse_to_insert = []
                    for sse in sse_after_insert:
                        if sse not in sse_before_insert:
                            sse_to_insert.append(sse)
                    
                    insert_data_into_db(sse_to_insert, 'A1788_software_server_environment', code)
                    
                    cse_to_insert = []
                    for cse in cse_after_insert:
                        if cse not in cse_before_insert:
                            cse_to_insert.append(cse)
                    
                    insert_data_into_db(cse_to_insert, 'A1788_computer_system_environment', code)
                    
                    print(f"{remaining} apps remaining.")
                    print("-------------------\n")
                    remaining -= 1
                break  # Task completed successfully, exiting the loop
            except Exception as e:
                print(f"An error occurred: {e}")
                close_db_connection()
                retry_count += 1
                if retry_count < MAX_RETRIES:
                    print("Retrying in 5 seconds...")
                else:
                    print("Maximum number of retries reached. Exiting...")
                time.sleep(5) # Adding a delay before restarting the loop
                continue
            finally:
                print("Done!")
                close_db_connection()