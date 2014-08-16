### Contains functions for working with the PER datasets

import imp
mods = ['xlrd'#,
        #'madgetech.madgetech_2',
        #'madgetech.madgetech_2_tow',
        #'madgetech.madgetech_2_cs',
        #'madgetech.madgetech_2_rht'
        ]
for i in mods:
    try:imp.find_module(i)
    except ImportError:print 'You are missing the '+i+' module'
