# TODO: Rewrite fileExplorer and make is use json located on one folder where everyone can get things from
# TODO: create a stopwatch that starts and stop, and keeps adding time to selected project


# 

import tkinter as tk
from tkinter import *
import os
import xlrd
import xlwt


class fileExplorer17(Frame):

    file_path = os.path.expanduser(
        "~/Desktop/scripts/python/fileexplorer17/fe17.xlsx")

    def __init__(self, master=None):
        super().__init__(master)
        # self.pack()
        self.create_widgets()

    # this function will search for a json in the file path, if it doesn't exist, it will create it.
    # The json will have an array with variables. Insert
    def store_json_variables(array_vars):
        data = fileExplorer17.update_data()
        def checkifVarExists(variable):
            if variable is not None:
                return variable
            else:
                return "None"

        filepath = checkifVarExists(array_vars[0])
        action_desired = checkifVarExists(array_vars[1])
        main_object_name = array_vars[2]
        name_variable_to_find = array_vars[3]
        value_to_store = array_vars[4]
        try:
            value_inside_array_to_find = array_vars[5]
        except:
            pass

        # ^actions: store, retrieve, delete
        # Check if the JSON file exists
        if os.path.exists(filepath):
            # If it exists, open it in read mode and load its contents
            with open(filepath, 'r') as file: 
                data = json.load(file)
        else:
            # If it doesn't exist, create a new object with an empty array
            data = {'variables': []}
            # create data in the file
            with open(filepath, 'w') as file:
                json.dump(data, file)

        # Store the new data in the array
        if action_desired == "store":
            # check if name_variable_to_find exists, if not, create it inside the first array found in the Json
            if name_variable_to_find in data["variables"][0]:
                # if it exists, update the value
                # print break lines 
                # print ("\n"*1)
                # print (value_to_store)
                data["variables"][0][name_variable_to_find] = value_to_store
                # print ("\n"*1 )
                # print (data["variables"][0][name_variable_to_find])
                # print("after")
                # dump the file
                with open(filepath, 'w') as file:
                    json.dump(data, file)
                # print(data)
            else:
                # if it doesn't exist, create it
                data["variables"][0][name_variable_to_find] = value_to_store
                # dump the file
                with open(filepath, 'w') as file:
                    json.dump(data, file)
        elif action_desired == "retrieve_index_from_array":
            # find the index of the variable that matches the value inside an array
            for i in range(len(data["variables"][0][name_variable_to_find])):
                if data["variables"][0][name_variable_to_find][i] == value_inside_array_to_find:
                    return i
        elif action_desired == "retrieve_value_from_array":
            # find the index of the variable that matches the value inside an array
            for i in range(len(data["variables"][0][name_variable_to_find])):
                if data["variables"][0][name_variable_to_find][i] == value_inside_array_to_find:
                    return data["variables"][0][name_variable_to_find][i]
                # Look for arrays inside the "games" arra
        elif action_desired == "retrieve_single_value":
            # retrieve value that matches the variable name name_variable_to_find
            return data["variables"][0][name_variable_to_find]
        elif action_desired == "delete":
            # I need to 
            # delete the variable name_variable_to_find
            #^ main_object_name is an array here 
            # 0 - selected_item_games
            # 1 - selected_fav_proj
            # 2 - selected_index
            # 3 which listbox is deleting something
            selected_listbox = name_variable_to_find
            if selected_listbox == 1:
                # print("deleting game")
                selected_item_games = main_object_name[0]
                # print(data["variables"][0][selected_item_games])
                del data["main"][selected_item_games]

            
            
            elif selected_listbox == 2:
                # ^ section 11 - delete object from data
                selected_item_games = main_object_name[0]
                selected_fav_proj = main_object_name[1]
                del data["main"][selected_item_games][selected_fav_proj]


            elif selected_listbox == 3:
                selected_item_games = main_object_name[0]
                selected_fav_proj = main_object_name[1]
                selected_index = main_object_name[2]
                # ^ section 11 - delete object from data
                del data["main"][selected_item_games][selected_fav_proj][selected_index]
            
            # print(data["main"][selected_item_games][selected_fav_proj][selected_index])
            # print(data["main"][selected_item_games][selected_fav_proj][selected_index])
            # path_chosen = data_main['main'][selected_item_games][selected_fav_proj][selected_index]['path']

            # del data["main"][0][name_variable_to_find]
            # dump the file
            with open(filepath, 'w') as file:
                json.dump(data, file)

    def refresh_listboxes(listbox_selected):
        # fe17.listbox.delete(0, END)
        # fe17.listbox_fav_proj.delete(0, END)
        if listbox_selected == 1:
            # fe17.listbox_favorites.delete(0, END)
            listbox3_selected = fe17.listbox.curselection()[0]
            # print(listbox3_selected)
            fe17.listbox_favorites.delete(listbox3_selected)
            # newList = fileExplorer17.listbox_3_return_display_name()
            # print(newList)
            # if newList != []:
            #     sortedTitles = sorted(newList)
            #     for item in sortedTitles:
            #         print(item)
            #         fe17.listbox_favorites.insert('end', item)

        # add options 
        # fileExplorer17.update_listbox_items(fe17.listbox, "title")
        # fileExplorer17.update_listbox_items(fe17.listbox_fav_proj, "fav_")
        fileExplorer17.update_listbox_items(fe17.listbox_favorites, "listbox3_favorites")



    def openOneDriveOnline(path2, contextMenuName, contextOption, subContextOption):
        rootFolder = os.path.dirname(path2)
        rootFolder_for_app = os.path.basename(rootFolder)
        openFolder = os.path.basename(path2)
        Application().start(r'explorer.exe "{}"'.format(rootFolder))
        '''
        # In Windows 11 ContextMenu is called Dialog

        '''
        # connect to another process spawned by explorer.exe
        # Note: make sure the script is running as Administrator!
        app = Application(backend="uia").connect(
            path="explorer.exe", title=rootFolder)
        # for k,v in app.items():
        #     print(v)
        # print(f'dd {rootFolder}')
        # print(f'df {openFolder}')
        app[rootFolder_for_app].set_focus()
        app[rootFolder_for_app].maximize()
        sleep(3)
        chosen_Folder_ctrl = app[rootFolder_for_app].ItemsView.get_item(
            openFolder)
        chosen_Folder_ctrl.right_click_input()
        # chosen_Folder_ctrl.print_control_identifiers()
        try:
            app['ContextMenu'][subContextOption].invoke()
        except:
            app['ContextMenu'][subContextOption].click_input()

        # if contextMenuName == "Dialog":
        #     app[contextMenuName][contextOption].click_input()
        #     sleep(1)
        #     app[contextMenuName][subContextOption].invoke()
        # if contextMenuName == "ContextMenu":
        #     app[contextMenuName][contextOption].invoke()

        # function to add input to first empty cell in first row of Excel file

    def find_elements(data, json_sub_locator, json_locator, json_name, isFound, isSubLocFound):
        # find how to when not finding the right title, move on
        """
        Finds all values for the given json_sub_locator in the data dictionary or list that has a matching 'title' value.

    Args:
            data (dict or list): The data to search for the json_sub_locator.
            json_sub_locator (str): The json_sub_locator to search for.
            title (str): The title of the object to search for.

        Returns:
           List with all objects accessible by elements[n]['name']
        """
        elements = []

        if isinstance(data, dict):
            # K is the name of the element, V is the content
            if isFound == True:
                for k, v in data.items():

                    if isinstance(v, (dict, list)):
                        '''
                        If the path object (eg favorites) has been found
                        '''
                        if k == json_sub_locator:

                            # If array is not empty
                            elements += fileExplorer17.find_elements(
                                v, json_sub_locator, json_locator, json_name, True, True)
                            break
                        else:
                            continue
            else:
                for k, v in data.items():

                    # ! if title doesn't match the one found, continue
                    if isinstance(v, (dict, list)):
                        # If array is not empty
                        if len(v) > 0:
                            elements += fileExplorer17.find_elements(
                                v, json_sub_locator, json_locator, json_name, False, False)
                        else:
                            continue
        #! If its an array
        elif isinstance(data, list):

            # Found
            if isFound == True:
                if isSubLocFound == True:
                    '''
                    This is where the function ends

                    Since it's a array already, no need to append
                    '''
                    elements = data
                else:
                    for item in data:
                        # Here you need to search further inside

                        elements += fileExplorer17.find_elements(
                            item, json_sub_locator, json_locator, json_name, True, False)
            else:
                for item in data:
                    # !This is where you need to search through titles to find
                    # the desired title
                    if isinstance(item, (dict, list)):

                        '''
                        Search through each title on the first
                        array until you find what you need,
                        then activate the bool
                        '''
                        if str.lower(item['title']) != str.lower(json_name):
                            continue
                        else:

                            elements += fileExplorer17.find_elements(
                                item, json_sub_locator, json_locator, json_name, True, False)
            # returns array
        return elements

    def get_game_arrays(file_path):
        desktop = os.path.join(os.path.join(
            os.environ['USERPROFILE']), 'Desktop')
        # file_path = os.path.join(desktop, 'scripts', 'data.json')

        # Check if the JSON file exists
        if os.path.exists(file_path):
            # If it exists, open it in read mode and load its contents
            with open(file_path, 'r') as file:
                data = json.load(file)

                # Look for arrays inside the "games" array
                game_arrays = []
                for item in data['games']:
                    if isinstance(item, list):
                        # If it's an array, create a new object with variable names for each of its elements
                        obj = {}
                        for i, value in enumerate(item):
                            obj['value{}'.format(i)] = value
                        game_arrays.append(obj)

                # Return the list of objects representing the arrays inside the "games" array
                return game_arrays
        else:
            # If the file doesn't exist, return an empty list
            return []

    def get_titles_from_json(file_location, value_name):
        # Define the file path for the JSON file

        # Check if the file exists
        if not os.path.isfile(file_location):
            print("The file does not exist!")
            return

        # Load the data from the JSON file
        with open(file_location, "r") as f:
            data = json.load(f)

        # Check if the "main" array exists in the data
        if "main" in data:
            # If it exists, get all the "title" elements from the array
            titles = [elem[value_name]
                      for elem in data["main"] if value_name in elem]

            return titles
            # return sorted(titles)
        else:
            print("The 'main' array does not exist in the data! from get titles")
            return []

    def listbox_return_index_value(listbox):
        # Get the selected item from the listbox
        selected_item = listbox.curselection()
        # Check if an item is selected
        if selected_item:
            # Get the index of the selected item
            index = selected_item[0]
            # Get the value of the selected item
            value = listbox.get(index)
        return value

    def add_to_favorites(listbox, input_fieldName_favorites, input_fieldName_favorites_display_name, locationToAddTo, createProjectBool = None):

        # Load the data from the JSON file
        with open(path_json, "r") as f:
            data = json.load(f)

        value = fileExplorer17.listbox_return_index_value(listbox)
        if locationToAddTo == "fav_folder":
            
            value_fav_folder_listbox = fileExplorer17.listbox_return_index_value(
                fe17.listbox_fav_proj)
        target_object = next(
            (obj for obj in data["main"] if obj["title"] == value), None)
        if target_object:
            # Check if the "favorites" array exists
            
            if createProjectBool == True:
                finalPath = input_fieldName_favorites
            else:
                finalPath = input_fieldName_favorites.get()
            # ^ section 9
            if finalPath.startswith('"') and finalPath.endswith('"'):
                # Remove double quotes
                finalPath = finalPath[1:-1]
            if "favorites" in target_object:
                # ^ not all function callings have this as a widget
                try:
                    display = input_fieldName_favorites_display_name.get()
                except:
                    display = input_fieldName_favorites_display_name
                if display == "":
                    if finalPath.startswith("http") or finalPath.startswith("www"):
                        display_name_new = "url_" + display
                    else:
                        display_name_new = os.path.basename(Path(finalPath))
                    # print(display_name_new)
                    new_object = {"path": os.path.normpath(
                        finalPath), "display_name": display_name_new}

                else:
                    if finalPath.startswith("http") or finalPath.startswith("www"):
                        display_name_new = "url_" + display
                        new_object = {"path": os.path.normpath(
                            finalPath), "display_name": display_name_new}

                    else:
                        # Create a new object with "path" and "display_name" values
                        new_object = {"path": os.path.normpath(
                            finalPath), "display_name": display}
                # Add the new object to the "favorites" array
                if locationToAddTo == "favorites":
                    target_object["favorites"].append(new_object)
                if locationToAddTo == "fav_folder":
                    target_object[value_fav_folder_listbox].append(new_object)
            else:
                # If the "favorites" array doesn't exist, create it and add the new object
                try:
                    target_object["favorites"] = [{"path": os.path.normpath(finalPath), "display_name": input_fieldName_favorites_display_name.get()}]
                except:
                    target_object["favorites"] = [{"path": os.path.normpath(finalPath), "display_name": input_fieldName_favorites_display_name}]
        else:
            # If no object is found with the matching title, raise an error
            raise ValueError(
                f"No object with title '{value}' found in the JSON data.")
        with open(path_json, "w") as f:
            json.dump(data, f)

        '''
        steps here:

        run vars to obtain arrays with all objects inside of each
        favorite, recording, and so on

        Change where to add depending on what button is calling it

        '''
        # Do something with the selected item
        all_favorites = fileExplorer17.find_elements(
            data_main, 'favorites', 'title', value, False, False)
        favorite_display_list = fileExplorer17.return_display_name_from_object(
            all_favorites, "display_name")
        # ^ section 8 - where new favorites are added to the listbox depending if the box is populated or empty
        try:
            display_name_input = input_fieldName_favorites_display_name.get()
        except:
            display_name_input = input_fieldName_favorites_display_name
        
        if display_name_input == "":
            if locationToAddTo == "favorites":
                fe17.listbox_main_fav.insert(END, display_name_new)

            if locationToAddTo == "fav_folder":
                fe17.listbox_favorites.insert(END, display_name_new)
        else:
            display_name = display_name_input
            if finalPath.startswith("http") or finalPath.startswith("www"):
                display_name2 = "url_" + display_name
                if locationToAddTo == "favorites":
                    fe17.listbox_main_fav.insert(
                        END, display_name2)
                if locationToAddTo == "fav_folder":
                    fe17.listbox_favorites.insert(
                        END, display_name2)
            else:
                if locationToAddTo == "favorites":
                    fe17.listbox_main_fav.insert(
                        END, display_name)
                if locationToAddTo == "fav_folder":
                    fe17.listbox_favorites.insert(
                        END, display_name)

        # Get the selected item

# Define a function to update the items in the listbox
    def update_listbox_items_fav(listbox_generic, new_items):
        # Delete all existing items from the listbox_generic
        listbox_generic.delete(0, END)

        if "main" in new_items:
            # If it exists, get all the "title" elements from the array
            titles = [elem["favorites"]["display_name"]
                      for elem in new_items["main"] if elem["favorites"]["display_name"] in elem]
        else:
            print("The 'main' array does not exist in the data!")

        # Insert the new items into the listbox_generic
        for item in titles:
            listbox_generic.insert(END, item)


    def copy_folder_and_rename_ppproj_and_aep_files_on_it(pathToFolderToCopy,pathtoPasteFolder,newFolderName):
        # get pathToFileToCopy and copy` it to pathTiPasteFolder with newFolderName
        # split string by spaces and add _ between words
        shutil.copytree(pathToFolderToCopy, pathtoPasteFolder + "/"+ newFolderName)
        # copy file from path_ratings17 to the folder "project files" inside newFolderName
        # shutil.copy(path_ratings17_aep_file, pathtoPasteFolder + "/" + newFolderName + "/project files")
    #  on new folder, search for .ppproj and .aep files and rename them to newFolderName
        selected_game = str(fileExplorer17.store_json_variables(
                [path_json_other_vars, "retrieve_single_value", "variables", "current_selection_games_name", "", ""]))
        
        createdate = datetime.now().strftime("%Y_%m_%d")
        aepFileName = "00_" + selected_game + "_" + createdate + "_" + newFolderName +"_Miguel" "_v01" 
        aepPMarketingFileName = aepFileName + "_Performance_Marketing_v01"
        prProjReversioning = aepFileName + "_Re-Versioning_v01"


        for root, dirs, files in os.walk(pathtoPasteFolder + "/" + newFolderName):
                    # ^Thunderful
                    # if filename is inside folder "After_Effects"
            currentFolder = root.split("\\")[-1]

            if currentFolder == "After_Effects":
                for filename in files:
                    if filename =="00_[Game]_YY.MM.DD_[Project_Name]_v01.aep":
                        if not os.path.exists(os.path.join(root, aepFileName + ".aep")):
                            os.rename(os.path.join(root, filename), os.path.join(root, aepFileName + ".aep"))
                    if filename =="00_Performance_Marketing_Template_v01.aep":
                        if not os.path.exists(os.path.join(root, aepPMarketingFileName + ".aep")):
                            os.rename(os.path.join(root, filename), os.path.join(root, aepPMarketingFileName + ".aep"))
            if currentFolder == "Premiere_Pro":
                for filename in files:
                    if filename == "00_[Game]_YY.MM.DD_[Project_Name]_v01.prproj":
                        if not os.path.exists(os.path.join(root, aepFileName + ".prproj")):
                            os.rename(os.path.join(root, filename), os.path.join(root, aepFileName + ".prproj"))
                    if filename == "00_Re-Versioning_Template.prproj":
                    
                        if not os.path.exists(os.path.join(root, prProjReversioning + ".prproj")):                    
                            os.rename(os.path.join(root, filename), os.path.join(root, prProjReversioning + ".prproj"))
                    
                    
                # ^T17
                # if filename.endswith(".prproj"):
                #     if not os.path.exists(os.path.join(root, newFolderName + ".prproj")):
                #         os.rename(os.path.join(root, filename), os.path.join(root, prProjFileName + ".prproj"))
                # if filename.endswith(".aep") and "ratings17" not in filename:
                #     if not os.path.exists (os.path.join(root, newFolderName + ".aep")):
                #         os.rename(os.path.join(root, filename), os.path.join(root, aepFileName + ".aep"))



    def listbox_3_return_display_name():
        # get listbox_fave selection
        data = fileExplorer17.update_data()
        listbox2selected = fe17.listbox_fav_proj.get(fe17.listbox_fav_proj.curselection())
        # search for all display_name inside of it
        all_display_names =[]
        for item in data['main']:
            if str.lower(item['title']) == str.lower(fe17.listbox.get(fe17.listbox.curselection())):
                for key, value in item.items():
                    if key.startswith("fav_"):
                        if key == listbox2selected:
                            for current_object in value:
                                for k,v in current_object.items():
                                    if k == "display_name":
                                        all_display_names.append(v)
        return all_display_names
        # return in list


    def create_json_object(file_name, file_location, optionChosen):
        # Define the file path for the JSON file
        file_path = os.path.expanduser("~/Desktop/example.json")

        # Check if the file exists
        if not os.path.isfile(file_location):
            print("The file does not exist!")
            return

        # Load the data from the JSON file
        with open(file_location, "r") as f:
            data = json.load(f)
        '''
        If you're adding a new game
        '''
        if fe17.clicked0.get() == optionChosen[0]:
            # Create the object with the specified properties
            new_object = {
                'title': file_name,
                'personal_drive': [],
                'onedrive': [],
                'favorites': [],
                'recording_drive': []
            }

            # Check if the "main" array exists in the data
            if "main" in data:
                # If it exists, add the new object to the array
                data["main"].append(new_object)
            else:
                # If it doesn't exist, create the "main" array and add the new object to it
                data["main"] = [new_object]

            # Write the updated data to the JSON file
            with open(file_location, "w") as f:
                json.dump(data, f)

            fileExplorer17.update_listbox_items(fe17.listbox, "title")
            '''
            If you're adding a new favorite folder
            '''
        elif fe17.clicked0.get() == optionChosen[1]:
            newObjectname = "fav_"+file_name
            # Create the object with the specified properties
            selection_games_listbox = fileExplorer17.listbox_return_index_value(
                fe17.listbox)

            # next returns the first item found when the condition is met
            index = next((i for i, item in enumerate(
                data['main']) if item['title'] == selection_games_listbox), None)
            target_object = next(
                (obj for obj in data["main"] if obj["title"] == selection_games_listbox), None)

            if target_object is not None:
                target_object[newObjectname] = []
            # Write the updated data to the JSON file
            with open(file_location, "w") as f:
                json.dump(data, f)

            fileExplorer17.update_listbox_items(fe17.listbox_fav_proj, "fav_")

            # print(fe17.clicked0.get())
        elif fe17.clicked0.get() == optionChosen[2]:
            # ^ Adds folder to fav project
            newObjectname = "fav_"+file_name
            # Create the object with the specified properties
            selection_games_listbox = fileExplorer17.listbox_return_index_value(
                fe17.listbox)

            # next returns the first item found when the condition is met
            index = next((i for i, item in enumerate(
                data['main']) if item['title'] == selection_games_listbox), None)
            target_object = next(
                (obj for obj in data["main"] if obj["title"] == selection_games_listbox), None)

            if target_object is not None:
                target_object[newObjectname] = []
            # Write the updated data to the JSON file
            with open(file_location, "w") as f:
                json.dump(data, f)

            fileExplorer17.update_listbox_items(fe17.listbox_fav_proj, "fav_")

            # ^ section - create folder inside project
            fileExplorer17.copy_folder_and_rename_ppproj_and_aep_files_on_it(path_vanilla_project_folder, pathProjectCreation, file_name)
            
            # fe17.listbox_fav_proj.selection_set(END)

            for index, value in enumerate(fe17.listbox_fav_proj.get(0, END)):
                if str.lower(value) == str.lower(newObjectname):
                    fe17.listbox_fav_proj.selection_set(index)


            # ^Adds folder to favorites
            fileExplorer17.add_to_favorites(fe17.listbox, os.path.join(pathProjectCreation, file_name), file_name,"fav_folder", True) 
            # fileExplorer17.add_to_favorites(self.listbox, "path that you get from created folder above", "name of path from created above_folder", "created_project_folder")) 
            
        # Define a function to update the items in the listbox

    def return_key_locators_with_x_on_it(json_key):
        data = fileExplorer17.update_data()
        titles = []
        for item in data['main']:
            if str.lower(item['title']) == str.lower(fe17.listbox.get(fe17.listbox.curselection())):
                for key, value in item.items():
                    if key.startswith(json_key):
                        titles.append(key)
            # print(titles)
        return titles
    

    def update_listbox_items(listbox_generic, json_key):

        data = fileExplorer17.update_data()
        # Delete all existing items from the listbox_generic
        listbox_generic.delete(0, END)
        if json_key == "title":
            if "main" in data:
                # If it exists, get all the "title" elements from the array
                titles = [elem["title"]
                          for elem in data["main"] if "title" in elem]
            else:
                print("The 'main' array does not exist in the data!")
        # ^ section 6.2
        elif json_key == "fav_":
            titles = fileExplorer17.return_key_locators_with_x_on_it("fav_")
            # ^Store current selected listbox item
            fileExplorer17.store_json_variables(
                [path_json_other_vars, "store", "variables", "current_listbox_with_pop_menu", str(listbox_generic)])
            fileExplorer17.store_json_variables(
                [path_json_other_vars, "store", "variables", "unordered_fav_project_list", titles])
        elif json_key == "listbox3_favorites":
            listbox2selected = fe17.listbox_fav_proj.get(fe17.listbox_fav_proj.curselection())
            listbox3selected = fe17.listbox_favorites.get(fe17.listbox_favorites.curselection())
            selected_fav_proj = fileExplorer17.store_json_variables([path_json_other_vars, "retrieve_value_from_array", "variables", "unordered_fav_project_list", "", listbox2selected])
            all_favorites = fileExplorer17.find_elements(data, 'favorites', 'title', selected_fav_proj, False, False)
            # print(all_favorites)
            titles = fileExplorer17.return_display_name_from_object(all_favorites, "display_name")

            # 
        sortedTitles = sorted(titles)
        # print(titles)

        # Insert the new items into the listbox_generic

        for item in sortedTitles:
            listbox_generic.insert(END, item)

    def navigate_pc_folders(options, name_to_search, path_to_search, options2):
        names_of_folders = []
        selected_item_games = fe17.listbox.get(fe17.listbox.curselection())
        if options == "names":
            for f in os.listdir(path_to_search):
                if str.lower(f.strip()) == str.lower(name_to_search.strip()) or f == "Desktop":
                    '''
                    Return all folders inside "Main game" in onedrive
                    '''
                    if options2 == "onedrive":
                        if os.path.isdir(os.path.join(path_to_search, selected_item_games, "Main Game")):
                            for t in os.listdir(os.path.join(path_to_search, selected_item_games, "Main Game")):
                                # print(os.path.join(path_to_search,selected_item_games,"Main Game"))
                                names_of_folders.append(t)
                        else:
                            for r in os.listdir(os.path.join(path_to_search, selected_item_games)):
                                names_of_folders.append(r)
                    else:
                        names_of_folders.append(f)
                # print(f)
            return names_of_folders
        if options == "return_path":
            '''
            Return all folders inside "Main game" in onedrive
            '''
                #^ from t17 structure
            # if options2 == "onedrive":
            #     if os.path.isdir(os.path.join(path_to_search, selected_item_games, "Main Game")):
            #         return os.path.join(path_to_search, selected_item_games, "Main Game", name_to_search)
            #         # for f in os.listdir(os.path.join(path_to_search,selected_item_games,"Main Game")):
            #     else:
            #         for r in os.listdir(os.path.join(path_to_search, selected_item_games)):
            #             if str.lower(r.strip()) == str.lower(name_to_search.strip()):
            #                 return os.path.join(path_to_search, selected_item_games, r)
            # ^from thunderful structure
            if options2 == "onedrive":
                if os.path.isdir(os.path.join(path_to_search, selected_item_games, "Brand Kit")):
                    return os.path.join(path_to_search, selected_item_games, "Brand Kit", name_to_search)
                    # for f in os.listdir(os.path.join(path_to_search,selected_item_games,"Main Game")):
                else:
                    for r in os.listdir(os.path.join(path_to_search, selected_item_games)):
                        if str.lower(r.strip()) == str.lower(name_to_search.strip()):
                            return os.path.join(path_to_search, selected_item_games, r)

            else:
                for f in os.listdir(path_to_search):
                    if str.lower(f.strip()) == str.lower(name_to_search.strip()):
                        return os.path.join(path_to_search, f)
            # fe17.listbox_recording_drive.insert(END, f)

    def on_select(event, switch):
        data_main = fileExplorer17.update_data()

        # Get the selected item
        selected_item = event.widget.get(event.widget.curselection())
        selected_item_index = event.widget.curselection()[0]
        # print(str(event.widget))
        '''
        steps here:

        run vars to obtain arrays with all objects inside of each 
        favorite, recording, and so on
        
        '''
        # print(type(event.widget))
        # ^Store current selected listbox item
        # print("\n"*1)
        # print(str(event.widget))
        fileExplorer17.store_json_variables(
            [path_json_other_vars, "store", "variables", "current_listbox_with_pop_menu", str(event.widget)])
        if switch == "listbox_games":

            fileExplorer17.store_json_variables(
                [path_json_other_vars, "store", "variables", "current_selection_games_name", selected_item])
            fileExplorer17.store_json_variables(
                [path_json_other_vars, "store", "variables", "current_selection_games_index", selected_item_index])
            fe17.listbox_fav_proj.delete(0, END)
            fe17.listbox_favorites.delete(0, END)
            fe17.listbox_onedrive.delete(0, END)
            fe17.listbox_main_fav.delete(0, END)
            if bool_hide_listboxes is not True:
                fe17.listbox_recording_drive.delete(0, END)
                fe17.listbox_n_drive.delete(0, END)
                fe17.listbox_u_drive.delete(0, END)

            # ^section 10
            # Do something with the selected item
            all_favorites = fileExplorer17.find_elements(
                data_main, 'favorites', 'title', selected_item, False, False)
            favorite_display_list = fileExplorer17.return_display_name_from_object(
                all_favorites, "display_name")

            for item in favorite_display_list:
                fe17.listbox_main_fav.insert(END, item)

            e_folders = fileExplorer17.navigate_pc_folders(
                "names", selected_item, path_recordings, "")
            for item in e_folders:
                fe17.listbox_recording_drive.insert(END, item)

            n_folders = fileExplorer17.navigate_pc_folders(
                "names", selected_item, path_n_drive, "")
            for item in n_folders:
                fe17.listbox_n_drive.insert(END, item)
            # u_folders = fileExplorer17.navigate_pc_folders(
            #     "names", selected_item, path_u_drive, "")
            # for item in u_folders:
            #     fe17.listbox_u_drive.insert(END, item)
            odrive_folders = fileExplorer17.navigate_pc_folders(
                "names", selected_item, path_onedrive, "onedrive")
            for item in odrive_folders:
                fe17.listbox_onedrive.insert(END, item)

            fileExplorer17.update_listbox_items(fe17.listbox_fav_proj, "fav_")
            # fileExplorer17.update_listbox_items(fe17.listbox_fav_proj,"fav_")
        elif switch == "listbox_fav":
            '''
            for the favorites inside folders
            '''
            fileExplorer17.store_json_variables(
                [path_json_other_vars, "store", "variables", "current_selection_fav_project_name", selected_item])
            fileExplorer17.store_json_variables(
                [path_json_other_vars, "store", "variables", "current_selection_fav_project_index", selected_item_index])

            currentGameSelected = fileExplorer17.listbox_return_index_value(
                fe17.listbox)
            fe17.listbox_favorites.delete(0, END)
            all_favorites = fileExplorer17.find_elements(
                data_main, selected_item, 'title', currentGameSelected, False, False)
            favorite_display_list = fileExplorer17.return_display_name_from_object(
                all_favorites, "display_name")
            # ^ section 7 - where the third top listbox is populated
            for item in favorite_display_list:
                fe17.listbox_favorites.insert(END, item)
                fileExplorer17.store_json_variables(
                    [path_json_other_vars, "store", "variables", "unordered_third_listbox_items", favorite_display_list])
            # ^ section 11 - stores third listbox values
            selected_game = str(fileExplorer17.store_json_variables(
                [path_json_other_vars, "retrieve_single_value", "variables", "current_selection_games_name", "", ""]))
            selected_fav_proj = str(fileExplorer17.store_json_variables(
                [path_json_other_vars, "retrieve_single_value", "variables", "current_selection_fav_project_name", "", ""]))
            selected_project_favorites = fileExplorer17.find_elements(
                data_main, selected_fav_proj, 'title', selected_game, False, False)
            # print(selected_project_favorites)
            favorite_path_list = fileExplorer17.return_display_name_from_object(
                selected_project_favorites, "path")

            fileExplorer17.store_json_variables(
                [path_json_other_vars, "store", "variables", "unordered_favorite_items_path_list", favorite_path_list])

        elif switch == "listbox_sub_fav":
            fileExplorer17.store_json_variables(
                [path_json_other_vars, "store", "variables", "current_selection_sub_fav_project_name", selected_item])
            fileExplorer17.store_json_variables(
                [path_json_other_vars, "store", "variables", "current_selection_sub_fav_project_index", selected_item_index])
            fileExplorer17.store_json_variables(
                [path_json_other_vars, "store", "variables", "current_selection_main_favorite", selected_item])
            pass

    def return_display_name_from_object(data, keyValue):
        display_names = []
        for item in data:
            if isinstance(item, dict):
                for k, v in item.items():
                    if k == keyValue:
                        display_names.append(v)
                    elif isinstance(v, (dict, list)):
                        display_names += fileExplorer17.return_display_name_from_object(
                            v)
            elif isinstance(item, list):
                display_names += fileExplorer17.return_display_name_from_object(
                    item)
        return display_names

    def update_data():
        with open(path_json, "r") as f:
            data_main = json.load(f)
        return data_main

    def explore(path):
        # explorer would choke on forward slashes
        path = os.path.normpath(path)

        if os.path.isdir(path):
            subprocess.run([FILEBROWSER_PATH, path])
        elif os.path.isfile(path):
            subprocess.run(
                [FILEBROWSER_PATH, '/select,', os.path.normpath(path)])

    def return_vanilla_index():
        van_index = ""
        for number, k in enumerate(data_main['main']):
            for r, v in k.items():
                if r == "title" and v == vanilla_files_json_name:
                    return number

    def clipboard_options(switch, value):
        if switch == "copy`":
            win32clipboard.OpenClipboard()
            data = win32clipboard.GetClipboardData()
            win32clipboard.CloseClipboard()
            finalpath = str("https:\\")+str(data)
            # set clipboard data
        else:
            win32clipboard.OpenClipboard()
            win32clipboard.EmptyClipboard()
            win32clipboard.SetClipboardText(value)
            win32clipboard.CloseClipboard()




    def open_path(event, switch, odrive_option=None):
        



        data_main = fileExplorer17.update_data()
        # Get the selected item from the listbox
        # print(event)
        # print(data_main['main'][selected_item_games]['favorites'])

        if switch == "copy_path_option":
            currentPath =fileExplorer17.get_last_active_listbox("path",event)
            # selected_index = event.curselection()[0]
            # ^section 14 - if path is empty, then it means that the user is trying to copy` a path from the main favorites listbox
            if currentPath == "":
                selected_item_ordered_list = fe17.listbox.get(fe17.listbox.curselection())
                selected_item__lb2_ordered_list = fe17.listbox_fav_proj.get(fe17.listbox_fav_proj.curselection())
                selected_item__lb3_ordered_list = fe17.listbox_favorites.get(fe17.listbox_favorites.curselection())
                selected_item_games = fileExplorer17.store_json_variables([path_json_other_vars, "retrieve_index_from_array", "variables", "unordered_games_list", "", selected_item_ordered_list])
                selected_fav_proj = fileExplorer17.store_json_variables([path_json_other_vars, "retrieve_value_from_array", "variables", "unordered_fav_project_list", "", selected_item__lb2_ordered_list])
                # selected_fav_proj = str(fileExplorer17.store_json_variables([path_json_other_vars, "retrieve_value_from_array", "variables", "unordered_fav_project_list", "", ]))
                selected_index = fileExplorer17.store_json_variables([path_json_other_vars, "retrieve_index_from_array", "variables", "unordered_third_listbox_items", "", selected_item__lb3_ordered_list])
                path_chosen = data_main['main'][selected_item_games][selected_fav_proj][selected_index]['path']
                # path_chosen = data_main['main'][selected_item_games]['favorites'][selected_index]['path']
            elif str(currentPath) == str(path_onedrive):
                current_selection = event.get(event.curselection())
                path_chosen = fileExplorer17.navigate_pc_folders("return_path", current_selection, path_onedrive, "onedrive")
            else:
                current_selection = event.get(event.curselection())
                path_chosen = fileExplorer17.navigate_pc_folders("return_path", current_selection, currentPath, "")
            fileExplorer17.clipboard_options("set", path_chosen)
            # path_chosen = fileExplorer17.navigate_pc_folders("return_path", selected_e_drive, path_recordings, "")
            return
            
        
        

        # ^Section 6.1
        #  I'll get name of selected game from  ordered list, then fetch the index of that name in the unordered list.
        selected_item_ordered_list = fe17.listbox.get(
            fe17.listbox.curselection())
        selected_item_games = fileExplorer17.store_json_variables(
            [path_json_other_vars, "retrieve_index_from_array", "variables", "unordered_games_list", "", selected_item_ordered_list])

        if switch == "main_favorites":
            # selected_item_games = fe17.listbox.curselection()[0]
            selected_index = event.widget.curselection()[0]
            path_chosen = data_main['main'][selected_item_games]['favorites'][selected_index]['path']
        elif switch == "sub_favorites":
            selected_index = event.widget.curselection()[0]
            # selected_fav_proj = fe17.listbox_fav_proj.get(fe17.listbox_fav_proj.curselection())
            ordered_list_selected_fav_proj = fe17.listbox_fav_proj.get(
                fe17.listbox_fav_proj.curselection())
            selected_fav_proj = str(fileExplorer17.store_json_variables(
                [path_json_other_vars, "retrieve_value_from_array", "variables", "unordered_fav_project_list", "", ordered_list_selected_fav_proj]))
            # selected_item_games = fe17.listbox.curselection()[0]

            path_chosen = data_main['main'][selected_item_games][selected_fav_proj][selected_index]['path']

        elif switch == "e_drive":
            selected_e_drive = fe17.listbox_recording_drive.get(
                fe17.listbox_recording_drive.curselection())
            path_chosen = fileExplorer17.navigate_pc_folders(
                "return_path", selected_e_drive, path_recordings, "")
        elif switch == "n_drive":
            selected_n_drive = fe17.listbox_n_drive.get(
                fe17.listbox_n_drive.curselection())
            path_chosen = fileExplorer17.navigate_pc_folders(
                "return_path", selected_n_drive, path_n_drive, "")
        elif switch == "o_drive":
            selected_o_drive = fe17.listbox_onedrive.get(
                fe17.listbox_onedrive.curselection())
            # if os.path.isdir(os.path.join)
            path_to_search = os.path.join(path_onedrive)
            # selected_item_games = fe17.listbox.get(fe17.listbox.curselection())
            path_chosen = fileExplorer17.navigate_pc_folders(
                "return_path", selected_o_drive, path_onedrive, "onedrive")
        elif switch == "vanilla_files":
            selected_index = event.widget.curselection()[0]

            van_index = fileExplorer17.return_vanilla_index()
            selected_van = fe17.listbox_vanilla_files.get(
                fe17.listbox_vanilla_files.curselection())
            path_chosen = data_main['main'][van_index]['favorites'][selected_index]['path']
        # print(path_chosen)
        if switch == "o_drive":
            # print("got here")
            if odrive_option != None and odrive_option == "just_open":
                fileExplorer17.explore(path_chosen)
            else:
                fileExplorer17.openOneDriveOnline(
                    path_chosen, "ContextMenu", "OneDrive")
        else:
            # check if starts with http, if so, open in browser
            if path_chosen.startswith("http"):
                # remove "url_" from the start of the path
                # path_chosen2 = path_chosen[4:]
                # webbrowser.get(r"C:\Program Files\Mozilla Firefox\firefox.exe").open_new_tab(path_chosen)
                webbrowser.get().open(path_chosen)
            else:
                fileExplorer17.explore(path_chosen)
        # fileExplorer17.clipboard_options("set",path_chosen)

    # ^ section 12 - get last active listbox
    def get_last_active_listbox(operation, event=None):
        
        active_listbox = fileExplorer17.store_json_variables([path_json_other_vars, "retrieve_single_value", "variables", "current_listbox_with_pop_menu", "", ""])
        # loop throuth array_listboxes and find the one matching active_listbox
        for i,value in enumerate(array_listboxes):
            if operation == "listbox":
                if str(array_listboxes[i][0]) == active_listbox:
                    # print(i)
                    return array_listboxes[i][0]
            else:
                if str(array_listboxes[i][0]) == active_listbox:
                    return array_listboxes[i][1]
        pass

    def menu_options(switch):
        '''
        This code works by using the open path function and using the same
        switch codes as the ones that open paths, with the exception that it
        turns a bool on that copies the path instead of opening it

        '''
        if switch == "opt_frame":
            win32clipboard.OpenClipboard()
            data = win32clipboard.GetClipboardData()
            win32clipboard.CloseClipboard()
            finalpath = "" + str("https://")+str(data)
            # print(finalpath)
            # set clipboard data
            win32clipboard.OpenClipboard()
            win32clipboard.EmptyClipboard()
            win32clipboard.SetClipboardText(finalpath)
            win32clipboard.CloseClipboard()
            # return
        if switch == "opt_copy":
            try:
                # copy_path_bool = True
                active_listbox = fileExplorer17.get_last_active_listbox("listbox")
                # print(active_listbox)
                fileExplorer17.open_path(active_listbox, "copy_path_option")
            except:
                pass
        if switch == "opt_onedrive_open":
            try:
                print("option")
                fileExplorer17.open_path(
                    fe17.listbox_onedrive, "o_drive", "view_online")
            except:
                pass
        if switch =="opt_remove":
            try:
                # get current selection from listbox
                active_listbox = fileExplorer17.get_last_active_listbox("listbox")
                # listbox1selected = fe17.listbox.curselection()[0]
                # listbox2selected = fe17.listbox_fav_proj.curselection()[0]
                # listbox3selected = fe17.listbox_favorites.curselection()[0]
                # print(listbox1selected,listbox2selected,listbox3selected)
                # get text from current selection
                # need to set the which listbox I need to delete from
                # get name of listbox from active_listbox
                # print(fe17.listbox)
                listbox_selected = 0
                if(active_listbox == fe17.listbox):
                    # print("tt")
                    listbox_selected = 1
                    listbox1selected = fe17.listbox.get(fe17.listbox.curselection())
                    # print("listbox1selected",listbox1selected)
                    selected_item_games = fileExplorer17.store_json_variables([path_json_other_vars, "retrieve_index_from_array", "variables", "unordered_games_list", "", listbox1selected])
                    # print("listbox1selected",listbox1selected)
                    fileExplorer17.store_json_variables([path_json, "delete", [selected_item_games], listbox_selected, "", ""])
                elif(active_listbox == fe17.listbox_fav_proj):
                    listbox_selected = 2
                    listbox1selected = fe17.listbox.get(fe17.listbox.curselection())
                    listbox2selected = fe17.listbox_fav_proj.get(fe17.listbox_fav_proj.curselection())
                    selected_item_games = fileExplorer17.store_json_variables([path_json_other_vars, "retrieve_index_from_array", "variables", "unordered_games_list", "", listbox1selected])
                    selected_fav_proj = fileExplorer17.store_json_variables([path_json_other_vars, "retrieve_value_from_array", "variables", "unordered_fav_project_list", "", listbox2selected])
                    fileExplorer17.store_json_variables([path_json, "delete", [selected_item_games,selected_fav_proj], listbox_selected, "", ""])
                elif(active_listbox == fe17.listbox_favorites):
                    listbox_selected = 3
                    listbox1selected = fe17.listbox.get(fe17.listbox.curselection())
                    listbox2selected = fe17.listbox_fav_proj.get(fe17.listbox_fav_proj.curselection())
                    listbox3selected = fe17.listbox_favorites.get(fe17.listbox_favorites.curselection())
                    selected_item_games = fileExplorer17.store_json_variables([path_json_other_vars, "retrieve_index_from_array", "variables", "unordered_games_list", "", listbox1selected])
                    selected_fav_proj = fileExplorer17.store_json_variables([path_json_other_vars, "retrieve_value_from_array", "variables", "unordered_fav_project_list", "", listbox2selected])
                    selected_index = fileExplorer17.store_json_variables([path_json_other_vars, "retrieve_index_from_array", "variables", "unordered_third_listbox_items", "", listbox3selected])
                    fileExplorer17.store_json_variables([path_json, "delete", [selected_item_games,selected_fav_proj,selected_index], listbox_selected, "", ""])
                # ^delete from json
                # ^refresh listboxes
                # fileExplorer17.refresh_listboxes(listbox_selected)
                    
            except:
                pass

    import json

    def copy_path_save_in_json(path_json, active_listbox, active_listbox_option):
        # Check if the json file exists, if not, create it
        try:
            with open(path_json) as f:
                data = json.load(f)
        except FileNotFoundError:
            data = {}
            with open(path_json, 'w') as f:
                json.dump(data, f)

        # Check if active_listbox is empty
        if not active_listbox:
            popup = tk.Tk()
            popup.wm_title("Error")
            label = tk.Label(
                popup, text="You need to select an option from the listbox where you tried to copy` the path")
            label.pack(side="top", pady=10)
            popup.after(1000, popup.destroy)
            popup.mainloop()
            return

        # Check if active_listbox_option is empty
        if not active_listbox_option:
            popup = tk.Tk()
            popup.wm_title("Error")
            label = tk.Label(
                popup, text="You need to select an option from the listbox option")
            label.pack(side="top", pady=10)
            popup.after(1000, popup.destroy)
            popup.mainloop()
            return

        # Update the json file with active_listbox and active_listbox_option
        data['current_listbox'] = active_listbox
        data['current_listbox_option'] = active_listbox_option
        with open(path_json, 'w') as f:
            json.dump(data, f)

    def create_json_file_if_not_exists(path, key, new_value):
        other_vars_path = os.path.join(path, "otherVars.json")
        if not os.path.exists(other_vars_path):
            data = {
                "main": [
                    {
                        "active_listbox": "test",
                        "path_1_folder": "test",
                        "path_2_folder": "test"
                    }
                ]
            }
            with open(other_vars_path, "w") as f:
                json.dump(data, f)
        with open(other_vars_path, "r") as f:
            data = json.load(f)
        for item in data["main"]:
            if key in item:
                item[key] = new_value
        with open(other_vars_path, "w") as f:
            json.dump(data, f)

    def do_popup(event, switch):
        try:
            # fileExplorer17.copy_path_save_in_json(path_json_other_vars,event.widget,switch)
            active_listbox = event.widget
            # print(active_listbox)
            active_listbox_option = switch
            fe17.m.tk_popup(event.x_root, event.y_root)
        finally:
            fe17.m.grab_release()

    def timetask_popup(event, switch):
        try:
            fe17.m_timetask.tk_popup(event.x_root, event.y_root)
        finally:
            fe17.m_timetask.grab_release()


    def create_widgets(self):
        import os
        self.current_frame = None
        var = IntVar()
        var.set("1")
        '''
        # ^Section 1
        self.frame_1 = LabelFrame(self, text="top frame", padx=20, pady=20)
        self.frame_1.grid(row=0, column=1)

        rowNoTimeClock = 0
        # ^lvl1
        # add textsaying "Project" to frame_1
        self.txt_fieldName = Label(
            self.frame_1, text="Project", pady=5).grid(row=rowNoTimeClock, column=0)



        # ^lvl2 
        rowNoTimeClock += 1
        # add entrybox to frame_1    
        # add button to frame_1


        # ^lvl3
        rowNoTimeClock += 1

        # add three listboxes to frame_1
        self.listbox_projects = Listbox(self.frame_1, width=20, height=10)
        self.listbox_projects.grid(row=rowNoTimeClock, column=0)

    # add right click menu to listbox
        self.listbox_projects.bind(
            "<Button-3>", lambda event: fileExplorer17.timetask_popup(event, "listbox"))

        self.listbox_sub_projects = Listbox(self.frame_1, width=20, height=10)
        self.listbox_sub_projects.grid(row=rowNoTimeClock, column=1)
        # self.btn_find_file = Button(self.frame_1, text="Find File", padx=20, pady=10, command=lambda: fileExplorer17.openOneDriveOnline(
        #     os.path.join(os.path.expanduser('~'), "OneDrive - Team17 Digital Limited\Products\Asset Requests"), "STOVE"))  # .grid(row=0,column=1)
        # self.btn_find_file.grid(row=0, column=1)
        self.m_timetask = Menu(self, tearoff=0)
        self.m_timetask.add_command(
            label="Copy Path", command=lambda: fileExplorer17.menu_options("opt_copy"))
        '''
        
        # ^Section 2 
        

        self.frame_2 = LabelFrame(self, text="Input area", padx=20, pady=20)
        self.frame_2.grid(row=1, column=1)
        self.clicked0 = StringVar()
        array_input_options = ["Add new Game", "Add new Project Folder","Add Json Project + Root Project Folder"]
        paths_to_save_project = str(fileExplorer17.store_json_variables([path_json_other_vars, "retrieve_single_value", "variables", "current_selection_games_name", "", ""]))

        self.clicked0.set(array_input_options[1])
        self.dd_input = OptionMenu(
            self.frame_2, self.clicked0, *array_input_options).grid(row=1, column=0)

        self.clicked0_1 = StringVar()
        self.clicked0_1.set(paths_to_save_project)
        # self.dd_input_2 = OptionMenu(
        #     self.frame_2, self.clicked0_1, *paths_to_save_project).grid(row=2, column=0)

        # self.txt_fieldName0 = Label(self.frame_2, text="Add Game",
        #                        pady=5).grid(row=1, column=0)

        self.input_fieldName0 = Entry(self.frame_2, border=2)
        self.input_fieldName0.grid(row=3, column=0)

        self.btn_add_to_worksheet = Button(
        self.frame_2, text="add to Json", padx=20, pady=10, command=lambda: fileExplorer17.create_json_object(self.input_fieldName0.get(), path_json, array_input_options))  # .grid(row=0,column=1)
        self.btn_add_to_worksheet.grid(row=4, column=0)

        '''
        # ^Section 2.1
        '''

        self.txt_fieldName_add_path1 = Label(self.frame_2, text="Add Item to Folder",
                                             pady=5).grid(row=1, column=3)

        self.input_fieldName_favorites1 = Entry(self.frame_2, border=3)
        self.input_fieldName_favorites1.grid(row=2, column=3)

        self.txt_fieldName_add_path_display_name1 = Label(self.frame_2, text="Display Name",
                                                          pady=5).grid(row=3, column=3)

        self.input_fieldName_favorites_display_name1 = Entry(
            self.frame_2, border=3)
        self.input_fieldName_favorites_display_name1.grid(row=4, column=3)

        self.btn_add_to_favorites1 = Button(
            self.frame_2, text="Add Item to Folder", padx=20, bg='#55a2d8', pady=10, command=lambda: fileExplorer17.add_to_favorites(self.listbox, self.input_fieldName_favorites1, self.input_fieldName_favorites_display_name1, "fav_folder"))  # .grid(row=0,column=1)
        self.btn_add_to_favorites1.grid(row=5, column=3)

        '''
        # ^Section 2.2
        '''

        self.txt_fieldName_add_path = Label(self.frame_2, text="Add Path to Default Items",
                                            pady=5).grid(row=1, column=4)

        self.input_fieldName_favorites = Entry(self.frame_2, border=4)
        self.input_fieldName_favorites.grid(row=2, column=4)

        self.txt_fieldName_add_path_display_name = Label(self.frame_2, text="Display Name",
                                                         pady=5).grid(row=3, column=4)

        self.input_fieldName_favorites_display_name = Entry(
            self.frame_2, border=4)
        self.input_fieldName_favorites_display_name.grid(row=4, column=4)

        self.btn_add_to_favorites = Button(
            self.frame_2, text="Add to Default Items", padx=20, bg="#55d89c", pady=10, command=lambda: fileExplorer17.add_to_favorites(self.listbox, self.input_fieldName_favorites, self.input_fieldName_favorites_display_name, "favorites"))  # .grid(row=0,column=1)
        self.btn_add_to_favorites.grid(row=5, column=4)

        '''
        #^ Section 3
        '''
        self.frame_3 = LabelFrame(self, text="Folder Area", padx=20, pady=20)
        self.frame_3.grid(row=2, column=1)

        self.txt_fieldName1 = Label(
            self.frame_3, text="Games", pady=5).grid(row=0, column=0)
        self.txt_fieldName2 = Label(self.frame_3, text="Project Folders",
                                    pady=5).grid(row=0, column=1)
        self.txt_fieldName3 = Label(self.frame_3, text="Item",
                                    pady=5).grid(row=0, column=2)
        if bool_hide_listboxes is not True:
            self.txt_fieldName4 = Label(self.frame_3, text="N Drive",
                                        pady=5).grid(row=0, column=3)
            self.txt_fieldName4 = Label(self.frame_3, text="U Drive",
                                        pady=5).grid(row=0, column=4)
            self.txt_fieldName5 = Label(self.frame_3, text="Recording Folder",
                                        pady=5).grid(row=0, column=5)
        self.txt_fieldName6 = Label(self.frame_3, text="OneDrive",
                                    pady=5).grid(row=0, column=6)
        self.txt_fieldName7 = Label(self.frame_3, text="Default Files",
                                    pady=5).grid(row=0, column=7)
        self.listbox = Listbox(self.frame_3, exportselection=False)

        # ^Section 3.1

        titles = fileExplorer17.get_titles_from_json(path_json, "title")
        fileExplorer17.store_json_variables(
            [path_json_other_vars, "store", "variables", "unordered_games_list", titles])

        self.listbox.grid(row=1, column=0)

        data = sorted(titles)

        for item in data:
            self.listbox.insert(END, item)

        
        # Bind the '<<ListboxSelect>>' event to the on_select function
        self.listbox.bind("<<ListboxSelect>>", lambda event: fileExplorer17.on_select(
            event, "listbox_games"))
        self.listbox.bind(
            '<FocusOut>', lambda e: fe17.listbox_favorites.selection_clear(0, END))
        self.listbox.bind(
            "<Button-3>", lambda event: fileExplorer17.do_popup(event, "sub_listbox1"))

        array_listboxes.append([self.listbox,""])
        # self.listbox.select_set(0)

        self.listbox.event_generate("<<ListboxSelect>>")
        self.listbox_fav_proj = Listbox(self.frame_3, exportselection=False)
        self.listbox_fav_proj.grid(row=1, column=1)
        self.listbox_fav_proj.bind("<<ListboxSelect>>", lambda event: fileExplorer17.on_select(
            event, "listbox_fav"))
        self.listbox_fav_proj.bind(
            "<Button-3>", lambda event: fileExplorer17.do_popup(event, "sub_listbox2"))

        array_listboxes.append([self.listbox_fav_proj,""])

        self.listbox_favorites = Listbox(self.frame_3, exportselection=False)
       # Bind the function to the listbox
        self.listbox_favorites.bind(
            "<<ListboxSelect>>", lambda event: fileExplorer17.on_select(event, "listbox_sub_fav"))
        self.listbox_favorites.bind(
            '<Double-1>', lambda event: fileExplorer17.open_path(event, "sub_favorites"))
        self.listbox_favorites.bind(
            "<Button-3>", lambda event: fileExplorer17.do_popup(event, "sub_favorites"))
        self.listbox_favorites.configure(
            background="#99ccf3", foreground="black")
        self.listbox_favorites.grid(row=1, column=2)
        array_listboxes.append([self.listbox_favorites,""])

        # ^ Hidden listboxes
        if bool_hide_listboxes is not True:
            self.listbox_n_drive = Listbox(self.frame_3, exportselection=False)
            self.listbox_n_drive.bind(
                "<<ListboxSelect>>", lambda event: fileExplorer17.on_select(event, "listbox_other"))
            self.listbox_n_drive.bind(
                '<Double-1>', lambda event: fileExplorer17.open_path(event, "n_drive"))
            self.listbox_n_drive.grid(row=1, column=3)
            array_listboxes.append([self.listbox_n_drive,path_n_drive])
            
            self.listbox_u_drive = Listbox(self.frame_3, exportselection=False)
            self.listbox_u_drive.bind(
                "<<ListboxSelect>>", lambda event: fileExplorer17.on_select(event, ""))
            self.listbox_u_drive.grid(row=1, column=4)
            array_listboxes.append([self.listbox_u_drive,path_u_drive])


            self.listbox_recording_drive = Listbox(
                self.frame_3, exportselection=False)
            self.listbox_recording_drive.bind(
                "<<ListboxSelect>>", lambda event: fileExplorer17.on_select(event, ""))
            self.listbox_recording_drive.bind(
                '<Double-1>', lambda event: fileExplorer17.open_path(event, "e_drive"))
            self.listbox_recording_drive.grid(row=1, column=5)
            array_listboxes.append([self.listbox_recording_drive,path_recordings])

        self.listbox_onedrive = Listbox(self.frame_3, exportselection=False)
        # print(self.listbox_onedrive)
        self.listbox_onedrive.bind(
            "<<ListboxSelect>>", lambda event: fileExplorer17.on_select(event, ""))
        self.listbox_onedrive.bind(
            '<Double-1>', lambda event: fileExplorer17.open_path(event, "o_drive", "just_open"))
        self.listbox_onedrive.bind(
            "<Button-3>", lambda event: fileExplorer17.do_popup(event, "o_drive"))
        self.listbox_onedrive.grid(row=1, column=6)
        array_listboxes.append([self.listbox_onedrive,path_onedrive])


        self.listbox_vanilla_files = Listbox(
            self.frame_3, exportselection=False)
        self.listbox_vanilla_files.bind(
            "<<ListboxSelect>>", lambda event: fileExplorer17.on_select(event, ""))
        self.listbox_vanilla_files.bind(
            '<Double-1>', lambda event: fileExplorer17.open_path(event, "vanilla_files"))
        self.listbox_vanilla_files.grid(row=1, column=7)
        self.listbox_vanilla_files.bind(
            "<Button-3>", lambda event: fileExplorer17.do_popup(event, "vanilla_files"))
        array_listboxes.append([self.listbox_vanilla_files,""])

        all_vanilla = fileExplorer17.find_elements(
            data_main, 'favorites', 'title', vanilla_files_json_name, False, False)
        vanilla_display_list = fileExplorer17.return_display_name_from_object(
            all_vanilla, "display_name")

        for item in vanilla_display_list:
            self.listbox_vanilla_files.insert(END, item)

        '''
        Section 4
        '''

        # self.frame_4 = LabelFrame(
        #     self, text="Sub Folder Area", padx=20, pady=20)
        # self.frame_4.grid(row=3, column=1)

        self.txt_fieldName1 = Label(
            self.frame_3, text="Default items", pady=5).grid(row=2, column=2
                                                             )

        self.listbox_main_fav = Listbox(self.frame_3, exportselection=False)
        self.listbox_main_fav.grid(row=3, column=2)
        self.listbox_main_fav.configure(
            background="#55d89c", foreground="black")
        self.listbox_main_fav.bind(
            '<Double-1>', lambda event: fileExplorer17.open_path(event, "main_favorites"))
        array_listboxes.append([self.listbox_main_fav,""])
        '''
        # ^Right click menu section 11
        '''
        # L = Label(self, text="Right-click to display menu", width=40, height=20)
        # L.pack()
        self.m = Menu(self, tearoff=0)
        self.m.add_command(
            label="Copy Path", command=lambda: fileExplorer17.menu_options("opt_copy"))
        self.m.add_command(label="Convert Frame.io link",
                           command=lambda: fileExplorer17.menu_options("opt_frame"))
        self.m.add_command(
            label="Odrive Online", command=lambda: fileExplorer17.menu_options("opt_onedrive_open"))
        self.m.add_command(label="Edit Display Name")
        self.m.add_separator()
        self.m.add_command(label="Remove", command=lambda: fileExplorer17.menu_options("opt_remove"))
        # L.bind("<Button-3>", do_popup)
    
    def copy_history():
        # &  spawn window in mouse location
        # mC = Menu(self, tearoff=0)
        # mC.add_command(
        #     label="Copy Path", command=lambda: fileExplorer17.menu_options("opt_copy"))
        print("d")
    #Do something



if __name__ == "__main__":
    import json
    import os
    from pathlib import Path
    import subprocess
    import win32clipboard
    from pywinauto import Desktop, Application
    import webbrowser
    import win32com.client
    from os import listdir
    from time import sleep
    from datetime import datetime
    import shutil
    import getpass
    import sys
    from time import sleep
    import keyboard
    import threading
    from time import sleep
    from threading import Event, Thread
    array_listboxes = []
    desktop = os.path.expanduser('~') + '/Desktop'
    user = os.path.expanduser('~')
    users = os.path.expanduser('~') + '/'
    # using f strings to get the username
    username = os.getlogin()
    # print(username)
    pathDefault = users + username + '/Desktop'
    paths_personal_laptop = [

        pathDefault,
        pathDefault,
        pathDefault,
        pathDefault,
        pathDefault,
        pathDefault,
        pathDefault,
        pathDefault,
        pathDefault 




    ]



    path_json = paths_personal_laptop[0]
    active_listbox = ""
    active_listbox_option = ""
    copy_path_bool = False
    path_recordings = paths_personal_laptop[1]
    path_n_drive =paths_personal_laptop[2]
    path_u_drive = paths_personal_laptop[3]
    path_ratings17_aep_file = paths_personal_laptop[4]
    username = os.getlogin()
    # path_onedrive = fr"C:\Users\{username}\OneDrive - Team17 Digital Limited\Products"
    # path_onedrive = r"H:\Shared drives\TG Marketing\[Titles and Franchises]"
    path_onedrive = paths_personal_laptop[5]
    bool_hide_listboxes = True
    vanilla_files_json_name = "Vanilla Files"
    path_vanilla_project_folder = paths_personal_laptop[6]
    path_json_other_vars = paths_personal_laptop[7]
    pathProjectCreation = paths_personal_laptop[8]
    # print(fileExplorer17.get_game_arrays(path_json))
    # Load the data from the JSON file
    with open(path_json, "r") as f:
        data_main = json.load(f)
    FILEBROWSER_PATH = os.path.join(os.getenv('WINDIR'), 'explorer.exe')

    # fileExplorer17.openOneDriveOnline(path7,"Dialog","OneDrive","View online")

    # create_json_file_if_not_exists(path, "active_listbox", "new_value")
    # set BROWSER=r"C:\Program Files\Mozilla Firefox\firefox.exe"



    # keyboard.wait("esc")
    # keyboard.add_hotkey('windows + v', print, args =('you entered', 'hotkey'))
    
    # 

    root = Tk()
    fe17 = fileExplorer17(root)
    fe17.pack()
    fe17.mainloop()


    