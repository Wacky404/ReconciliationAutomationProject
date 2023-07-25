from reconcile import DataFile as df

mississippi = df(r"C:\Users\Wayne Cole\Downloads\Work Stuff\Copy Mississippi Educational Institutions 2023-05-26.xlsx"
                 , "All Mississippi Institution T1", "MS")
mississippi.raw_file_check(mississippi.raw_file)

alabama = df(r"C:\Users\Wayne Cole\Downloads\Work Stuff\Copy Alabama Educational Institutions 2023-06-20.xlsx"
             , "All Alabama Institutions T1", "AL")
alabama.raw_file_check(alabama.raw_file)

# mississippi.reconcile_institution(mississippi.wb_uasys, mississippi.ws_uasys, mississippi.raw_file
#                                   , df.ws_data_grab, df.ws_nces_grab)
# mississippi.reconcile_governing(mississippi.wb_uasys, mississippi.ws_uasys, mississippi.raw_file
#                                 , mississippi.abbrev, df.ws_data_grab, df.ws_nces_grab)
# mississippi.reconcile_campuslocation(mississippi.wb_uasys, mississippi.ws_uasys, mississippi.raw_file
#                                      , mississippi.abbrev, df.ws_data_grab, df.ws_nces_grab)
#
# alabama.reconcile_institution(alabama.wb_uasys, alabama.ws_uasys, alabama.raw_file
#                               , df.ws_data_grab, df.ws_nces_grab)
# alabama.reconcile_governing(alabama.wb_uasys, alabama.ws_uasys, alabama.raw_file
#                             , alabama.abbrev, df.ws_data_grab, df.ws_nces_grab)
# alabama.reconcile_campuslocation(alabama.wb_uasys, alabama.ws_uasys, alabama.raw_file
#                                  , alabama.abbrev, df.ws_data_grab, df.ws_nces_grab)

print('Starting Mississippi....')
mississippi.ai_institution(mississippi.wb_uasys, mississippi.ws_uasys, mississippi.raw_file)
mississippi.ai_campuslocation(mississippi.wb_uasys, mississippi.ws_uasys, mississippi.raw_file)
print('Starting Alabama....')
alabama.ai_institution(alabama.wb_uasys, alabama.ws_uasys, alabama.raw_file)
alabama.ai_campuslocation(alabama.wb_uasys, alabama.ws_uasys, alabama.raw_file)
