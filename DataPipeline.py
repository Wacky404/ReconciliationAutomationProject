from reconcile import DataFile as df

amount = int(input('How many states are you reconciling/cleansing?(integer) '))
i = 0
while i < amount:
    state = []
    user_state = input('Name of state: ')
    state.append(user_state)

    file_location = []
    user_file_location = str(input('File location of workbook(.xlsx): '))
    file_location.append(user_file_location)

    worksheet = []
    user_worksheet = str(input('Exact worksheet name: '))
    worksheet.append(user_worksheet)

    abrev_state = []
    user_abrev_state = str(input('State abbreviation: '))
    abrev_state.append(user_abrev_state)
    i += 1
go = True
while go:
    user_choice = int(input('Do you want to Reconcile --> 1 | Reconcile+Cleanse --> 2 | Cleanse --> 3 '))
    if user_choice == 1:
        for i in range(len(state)):
            state[i] = df(file_location[i], worksheet[i], abrev_state[i])

            state[i].reconcile_institution(state[i].wb_uasys, state[i].ws_uasys, state[i].raw_file, df.ws_data_grab, df.ws_nces_grab)
            state[i].reconcile_governing(state[i].wb_uasys, state[i].ws_uasys, state[i].raw_file, state[i].abbrev, df.ws_data_grab, df.ws_nces_grab)
            state[i].reconcile_campuslocation(state[i].wb_uasys, state[i].ws_uasys, state[i].raw_file, state[i].abbrev, df.ws_data_grab, df.ws_nces_grab)
    elif user_choice == 2:
        for i in range(len(state)):
            state[i] = df(file_location[i], worksheet[i], abrev_state[i])

            state[i].reconcile_institution(state[i].wb_uasys, state[i].ws_uasys, state[i].raw_file, df.ws_data_grab, df.ws_nces_grab)
            state[i].reconcile_governing(state[i].wb_uasys, state[i].ws_uasys, state[i].raw_file, state[i].abbrev, df.ws_data_grab, df.ws_nces_grab)
            state[i].reconcile_campuslocation(state[i].wb_uasys, state[i].ws_uasys, state[i].raw_file, state[i].abbrev, df.ws_data_grab, df.ws_nces_grab)

            print('Reconcile is done moving on to cleaning....')

            state[i].clean_governing(state[i].wb_uasys, state[i].ws_uasys, state[i].raw_file)
            state[i].clean_institution(state[i].wb_uasys, state[i].ws_uasys, state[i].raw_file)
            state[i].clean_campuslocation(state[i].wb_uasys, state[i].ws_uasys, state[i].raw_file)
    elif user_choice == 3:
        for i in range(len(state)):
            state[i] = df(file_location[i], worksheet[i], abrev_state[i])

            state[i].clean_governing(state[i].wb_uasys, state[i].ws_uasys, state[i].raw_file)
            state[i].clean_institution(state[i].wb_uasys, state[i].ws_uasys, state[i].raw_file)
            state[i].clean_campuslocation(state[i].wb_uasys, state[i].ws_uasys, state[i].raw_file)
    else:
        print('You did not input any integer between 1 - 3, please try again')

    go = bool(input('Do you want to exit: Yes --> 0 | No --> 1'))
