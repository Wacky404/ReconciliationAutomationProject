from reconcile import DataFile as df

tennessee = df(r"C:\Users\Wayne Cole\Downloads\Work Stuff\Copy Tennessee Educational Institutions 2023-05-26.xlsx"
               , "All Tennessee Institutions", "TN")
tennessee.clean_governing(tennessee.wb_uasys, tennessee.ws_uasys, tennessee.raw_file)
tennessee.clean_institution(tennessee.wb_uasys, tennessee.ws_uasys, tennessee.raw_file)
tennessee.clean_campuslocation(tennessee.wb_uasys, tennessee.ws_uasys, tennessee.raw_file)

texas = df(r"C:\Users\Wayne Cole\Downloads\Work Stuff\Copy Texas Educational Institutions 2023-05-26.xlsx"
           , "All Texas Institutions", "TX")
texas.clean_governing(texas.wb_uasys, texas.ws_uasys, texas.raw_file)
texas.clean_institution(texas.wb_uasys, texas.ws_uasys, texas.raw_file)
texas.clean_campuslocation(texas.wb_uasys, texas.ws_uasys, texas.raw_file)

california = df(r"C:\Users\Wayne Cole\Downloads\Work Stuff\Copy California Educational Institutions 2023-06-20.xlsx"
                , "All California InstitutionsTest", "CA")
california.clean_governing(california.wb_uasys, california.ws_uasys, california.raw_file)
california.clean_institution(california.wb_uasys, california.ws_uasys, california.raw_file)
california.clean_campuslocation(california.wb_uasys, california.ws_uasys, california.raw_file)

mississippi = df(r"C:\Users\Wayne Cole\Downloads\Work Stuff\Copy Mississippi Educational Institutions 2023-05-26.xlsx"
                 , "All Mississippi Institution T1", "MS")
mississippi.clean_governing(mississippi.wb_uasys, mississippi.ws_uasys, mississippi.raw_file)
mississippi.clean_institution(mississippi.wb_uasys, mississippi.ws_uasys, mississippi.raw_file)
mississippi.clean_campuslocation(mississippi.wb_uasys, mississippi.ws_uasys, mississippi.raw_file)




