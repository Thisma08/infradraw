import win32com.client
import mysql.connector
from mysql.connector import Error
import sys
import os
import configparser
import time

MAX_RETRIES = 3

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
        conf_auth_plugin = config.get('Database', 'auth_plugin')
        
        connection = mysql.connector.connect(
            host = conf_host,
            database = conf_database,
            user = conf_user,
            password = conf_password,
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

def fetch_all_codes(connection):
    try:
        cursor = connection.cursor()
        cursor.execute("SELECT a_code FROM A1788_application")   
        result = cursor.fetchall()
        if result:
            return [row[0] for row in result]
        else:
            return []
    except Error as e:
        print(f"Error getting app codes: {e}")
        return None


def fetch_app_name(code, connection):
    try:
        cursor = connection.cursor()
        cursor.execute(f"SELECT name_fr FROM A1788_application WHERE a_code = '{code}'")   
        result = cursor.fetchone()
        if result:
            return result[0]
        else:
            return None
    except Error as e:
        print(f"Error getting application name: {e}")
        return None
    
def fetch_env_names(code, connection):
    try:
        cursor = connection.cursor()
        cursor.execute(f"""SELECT env_name 
                       FROM A1788_environment 
                       WHERE env_name LIKE '{code}%'""")    
        result = cursor.fetchall()
        if result:
            return [row[0] for row in result]
        else:
            return []
    except Error as e:
        print(f"Error getting environment names: {e}")
        return None

def fetch_env_type(env_name, connection):
    try:
        cursor = connection.cursor()  
        cursor.execute(f"""SELECT env_type 
                       FROM A1788_environment 
                       WHERE env_name = '{env_name}'""")   
        result = cursor.fetchone()
        if result:
            return result[0]
        else:
            return None
    except Error as e:
        print(f"Error getting environment types: {e}")
        return None
    
def fetch_cs_names(code, connection):
    try:
        cursor = connection.cursor() 
        cursor.execute(f"""SELECT DISTINCT cs.cs_name
                       FROM A1788_computer_system cs
                       LEFT JOIN A1788_software_server ss ON cs.id = ss.cs_id
                       LEFT JOIN A1788_software_server_environment sse ON ss.id = sse.ss_id
                       LEFT JOIN A1788_environment env ON sse.env_id = env.env_id
                       LEFT JOIN A1788_application app ON env.app_id = app.app_id AND app.a_code = '{code}'""") 
        result = cursor.fetchall()
        if result:
            return [row[0] for row in result]
        else:
            return []
    except Error as e:
        print(f"Error getting computer system names: {e}")
        return None

def fetch_cs_in_env(env_name, connection):
    try:
        cursor = connection.cursor() 
        cursor.execute(f"""SELECT cs.cs_name
                        FROM A1788_computer_system cs
                        JOIN A1788_computer_system_environment cse ON cs.id = cse.cs_id
                        JOIN A1788_environment env ON cse.env_id = env.env_id
                        WHERE env.env_name LIKE '{env_name}'""")  
        result = cursor.fetchall()
        if result:
            return [row[0] for row in result]
        else:
            return []
    except Error as e:
        print(f"Error getting computer systems in environment: {e}")
        return None

def fetch_ss_in_env(env_name, connection):
    try:
        cursor = connection.cursor() 
        cursor.execute(f"""SELECT ss.identification
                        FROM A1788_software_server ss
                        JOIN A1788_software_server_environment sse ON ss.id = sse.ss_id
                        JOIN A1788_environment env ON sse.env_id = env.env_id
                        WHERE env.env_name LIKE '{env_name}'""") 
        result = cursor.fetchall()
        if result:
            return [row[0] for row in result]
        else:
            return []
    except Error as e:
        print(f"Error getting software servers in environment: {e}")
        return None

def fetch_ss_linked_to_cs(cs_name, connection):
    try:
        cursor = connection.cursor()  
        cursor.execute(f"""SELECT ss.identification 
                       FROM A1788_software_server ss
                       JOIN A1788_computer_system cs ON ss.cs_id = cs.cs_id 
                       WHERE cs.cs_name LIKE '{cs_name}'""")   
        result = cursor.fetchall()
        if result:
            return [row[0] for row in result]
        else:
            return []
    except Error as e:
        print(f"Error getting software servers linked to computer systems: {e}")
        return None
    
def find_shape_with_text(page, visio, text_to_find):
    shapes = page.Shapes
    for shape in shapes:
        if shape.Characters.Text == text_to_find:
            return shape
    return None

def connect_shapes(visio, page, shape1, shape2):
    connector = page.Drop(visio.Application.ConnectorToolDataObject, 0.5, 0.5)
    connector.Cells("BeginX").GlueTo(shape1.Cells("PinX"))
    connector.Cells("EndX").GlueTo(shape2.Cells("PinX"))
    connector.CellsU("LineWeight").Formula = "3pt"

    return connector

def draw_server_shape(visio, cs_name, page, x_pos, y_pos):
    # Load the stencil file
    server_stencils_filename = "‪C:\\Program Files (x86)\\Microsoft Office\\root\\Office16\\Visio Content\\1036\\SERVER_M.VSSX"
 
    # 64 means “open and hidden”
    server_stencils = visio.Documents.OpenEx(server_stencils_filename, 64)
     
    server_stencil=server_stencils.Masters("Serveur")
    
    server = page.Drop(server_stencil, x_pos, y_pos)
    
    server.Text = cs_name
      
    return server

def create_visio_document(visio, app_codes):
    # Creating new Visio document
    for app_code in app_codes:
        print(app_code)
        print("------------")
        doc = visio.Documents.Add("")
   
        # Fetching app environments names
        environment_names = fetch_env_names(app_code, connection)
        
        if not environment_names:
            print("No environment names found.")
            sys.exit()
           
        # Creating an output folder
        folder_name = "c:\\exported_diagrams\\" + app_code
        if not os.path.exists(folder_name):
            os.makedirs(folder_name)
        
        environment_names.sort()
        print(environment_names)
        
        for env_name in environment_names:
            # Fetching Computer Systems names
            cs_names = fetch_cs_in_env(env_name, connection)
            cs_amount = len(cs_names) 
            
            if cs_amount > 0:
                # Creating a page
                page = doc.Pages.Add()
                activePage = page  # Storing reference to the active page
                activePage.Name = env_name
                
                env_desc_label = activePage.DrawRectangle(0, 11.49, 10, 10.99)
                app_name = fetch_app_name(app_code, connection)
                env_type = fetch_env_type(env_name, connection)
                env_desc_label.Text = f"Application {app_code} ({app_name}) : {env_name} ({env_type})"
                
                # Creating Computer Systems container and its label
                cont_cs_label = activePage.DrawRectangle(0.2, 10.49, 0.2, 9.99)
                cont_cs_label.Text = "COMPUTER SYSTEMS"
                
                cont_cs = activePage.DrawRectangle(0.2, 9.99, 0.2, 7.59)
                
                # Setting initial x position for the Computer Systems
                cs_x_pos = 2.9
                            
                cs_names.sort()
                
                # Drawing the Computer Systems
                for cs_name in cs_names:
                    cont_cs_label.Resize(0, 5.4, 65)
                    cont_cs.Resize(0, 5.4, 65)
                    server = draw_server_shape(visio, cs_name, activePage, cs_x_pos, 8.9)
                    if server:
                        server.Text = cs_name
                        # Repositioning the computer system's name
                        cell_to_edit = server.Cells("Controls.Y")
                        cell_to_edit.FormulaU = "TxtHeight*3.6"
                    cs_x_pos += 5.4
                                  
                cont_cs_width_mm = cont_cs.CellsU("Width").Formula
                cont_cs_width_mm = cont_cs_width_mm.replace(" mm", "")
                cont_cs_width_mm = cont_cs_width_mm.replace(",", ".")
                cont_cs_width_mm = float(cont_cs_width_mm)
                cont_cs_width_in = cont_cs_width_mm / 25.4
                
                # Creating Software Servers container and its label
                cont_ss_label = activePage.DrawRectangle(0.2, 6.39, 0.2 + cont_cs_width_in, 5.89)
                cont_ss_label.Text = "SOFTWARE SERVERS"
                
                cont_ss = activePage.DrawRectangle(0.2, 5.89, 0.2 + cont_cs_width_in, 5.89)
                
                ss_x_pos = 0.6
                ss_y_pos = 4.49
                
                longest_ss_column_length = 0
                
                for cs_name in cs_names:           
                    column_length = 0
                    ss_names = fetch_ss_linked_to_cs(cs_name, connection)
                    
                    # Handling duplicates
                    fixed_ss_names = []
                    for ss_name in ss_names:                 
                        if ss_name not in fixed_ss_names:
                            fixed_ss_names.append(ss_name)
                                            
                    fixed_ss_names.sort()
                    
                    if fixed_ss_names:
                        ss_rect = activePage.DrawRectangle(ss_x_pos - 0.2, ss_y_pos + 1.2, ss_x_pos + 5, ss_y_pos + 1.2)
                        # Connecting the two shapes
                        found_shape = find_shape_with_text(activePage, visio, cs_name)
                        if found_shape:
                            connect_shapes(visio, activePage, ss_rect, found_shape)
                        else:
                            print("Shape not found.")
                        
                    for ss_name in fixed_ss_names:
                        ss_rect.Resize(6, 1.4, 65)
                        ss = activePage.DrawRectangle(ss_x_pos, ss_y_pos, ss_x_pos + 4.8, ss_y_pos + 1)
                        ss.Text = ss_name
                        ss_y_pos -= 1.4
                        column_length += 1
                        
                    if column_length > longest_ss_column_length:
                        longest_ss_column_length = column_length
                                        
                    ss_x_pos += 5.4
                    ss_y_pos = 4.49
        
                cont_ss.Resize(6, (longest_ss_column_length * 1.4) + 0.4, 65)
                        
                cont_cs_label.Resize(0, 0.2, 65)
                cont_cs.Resize(0, 0.2, 65)
                cont_ss_label.Resize(0, 0.2, 65)
                cont_ss.Resize(0, 0.2, 65)
            
                # Resizing the page to fit all
                activePage.ResizeToFitContents()
            
                print(f"Diagram for {env_name} created successfully.")
            
                # Saving as a svg file
                filename = os.path.join(folder_name, f"{env_name}.svg")
                exportPath = os.path.abspath(filename)
                print(f"Diagram exported at location {exportPath} successfully.")
                activePage.Export(filename)
                                    
if __name__ == "__main__":
    retry_count = 0
    while retry_count < MAX_RETRIES:
        establish_db_connection()
        if connection is not None:
            try:
                visio = win32com.client.Dispatch("Visio.Application")
                visio.visible = False
                app_codes = fetch_all_codes(connection)
                create_visio_document(visio, app_codes)
                break  # Task completed successfully, exiting the loop
            except Exception as e:
                print(f"An error occurred: {e}")
                close_db_connection()
                retry_count += 1
                if retry_count < MAX_RETRIES:
                    print("Retrying in 5 seconds...")
                else:
                    print("Maximum number of retries reached. Exiting...")
                time.sleep(5)  # Adding a delay before restarting the loop
                continue
            finally:
                close_db_connection()
                # print(locals())
                if 'visio' in locals():
                    for doc in visio.Documents:
                        doc.Saved = True
                    visio.Quit()
        