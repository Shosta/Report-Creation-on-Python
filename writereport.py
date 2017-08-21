# coding: utf-8
"""
Write the eIoT report automatically based on the files in the directory tree.
"""

import sys
import getopt
import createdata
import SecurityReport
import writexlsxreport
import writepptreport


def usage():
    '''
    Display the awaited command and arguments for the 'writereport.py'.
    '''
    print('Here is the default usage of that command:')
    print('writereport.py -s <security_projects_folder>')
    print('If no source folder is specified, the default source folder is the '
          'current folder.')


def __get_security_projects_folder(argv):
    '''
    Get the Security Project folder from the command argument.
    The Security Project folder is the folder that contains all the information about the objects
    that are stored in subfolders.

    Params:
    argv the command arguments

    Return: The value that comes after the -s or --security_projects_folder command argument.
    '''
    import variables
    # Add default value.
    security_projects_folder = variables.SECURITY_PROJECTS_DIRPATH
    try:
        opts, args = getopt.getopt(argv, "hs:", ["help=", "security_projects_folder="])
    except getopt.GetoptError:
        print('Type \'writereport.py -h\' for help.')
        sys.exit(2)

    if opts.__len__() == 0:
        usage()
        #raise Exception('Usage displayed, stop application')

    for opt, arg in opts:
        if opt in ("-h", "--help"):
            usage()
            #raise Exception('Usage displayed, stop application')
        elif opt in ("-s", "--security_projects_folder"):
            security_projects_folder = arg
            variables.SECURITY_PROJECTS_DIRPATH = arg

    return security_projects_folder


def main(argv):
    """The main function that is launched when the program is started."""
    # Get folder argument from the command.
    try:
        security_projects_folder = __get_security_projects_folder(argv)
    except Exception:
        sys.exit()
        print('exit')

    # Populate the IoTObjects list with the info from the Security Projects folder.
    iot_objects_array = createdata.populate_objects_array()

     # Create the security report.
    security_report = SecurityReport.SecurityReport()
    security_report.populate_member_values(iot_objects_array)

    # Write the Security report.
    writexlsxreport.write_excel_file_from(
        security_projects_folder,
        'Rapport de Suivi de Sécurité des Objets Connectés.xlsx',
        iot_objects_array)

    writepptreport.write_dashboard_file(
        security_projects_folder,
        'Dashboard de Suivi de Sécurité des Objets Connectés.pptx',
        security_report)

    # Save the current SecurityReport to an xml file.
    security_report.save_to_file()

    print("Reports writing done")


if __name__ == '__main__':
    main(sys.argv[1:])
