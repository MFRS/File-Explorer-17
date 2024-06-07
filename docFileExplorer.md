# store_json_variables

Obtain info from listboxes stored in OtherVars.json

# Task time list

## Implementation

There will be a button on the top of the fileExplorer which will spawn the ui, easier as there's hardly any space for extra ui.

The idea is to have three listboxes. One will have the projects, all sorted from the start, another will have sub tasks , and the other will have the time ticking.

Since it's possible I won't be so granular on each task, I'll add a project element everytime I add a new project to time.

## Options

All options are tucked in in right clicking on the listboxes, which will allow me to pop_up entrys with one box.

## Further Ideas

Button to export report of time of all tasks or single task

Use rowtime and add 1 on each level to not have to manually change each one because of this.



# ^ Code Logic:

    # ^ create_widgets function

    # ^ section 1 - open button top

    # ^ section 2 - Add project or project folder section

    # ^ section 2.1 - Add path to project section

    # ^ section 2.2 - Add path to Default Paths section

    # ^ section 3 - Listbox Section

    # ^ section 3.1 - Lists games when boots up

    # ^ on_select lists all items on the listbox.

    # ^ open_path is where the double click function is.

    # ^ Section 6.1 -   I'll get name of selected game from  ordered list, then fetch the index of that name in the unordered list.

    # ^ selected_item_games is the var where I will assign the index to the matching name in the unordered value.

    # ^section 6.2 - Stores list of fav_proects in json

    # ^ add_to_favorites contains the other area where new values are added

    # ^ section 7 -  where the third top listbox is populated and it's stored in json

    # ^ section 8 - where new favorites are added to the listbox depending if the box is populated or empty. items are added to listbox

    # ^ section 9 is where paths and urls are added to favorites section. items are added to json

    # ^ section 10 is where the lists are populated when you look for a game name

    # ^section 11 right click menu

    #    ^ def menu_options is where the right click menu options are

    #  ^ opt_copy is the copy` path option

    # ^ section 12 - get last active listbox

    # ^ section 13 - return path from selected list

    # ^ section 14 - create folder inside project

    # & refresh option

    # * listbox_fav_proj is the second listbox in the top

    # * listbox_favorites is the third listbox in the top
