"""
Write the eIoT report automatically based on the files in the directory tree.
"""

import createdata
import writexlsxreport
import writepptreport


def main():
    """The main function that is launched when the program is started."""
    objects_dictionnary_array = createdata.populate_projects_array()

    writexlsxreport.write_excel_file_from(
        'Rapport de Suivi de Sécurité des Objets Connectés.xlsx',
        objects_dictionnary_array)

    writepptreport.write_ppt(
        'Dashboard de Suivi de Sécurité des Objets Connectés.pptx')

main()
print("Done")
