# coding: utf-8
"""
Write the eIoT report automatically based on the files in the directory tree.
"""

import createdata
import SecurityReport
import writexlsxreport
import writepptreport
from report_creation_utils import xml_utils


def main():
    """The main function that is launched when the program is started."""

    iot_objects_array = createdata.populate_objects_array()

    writexlsxreport.write_excel_file_from(
        'Rapport de Suivi de Sécurité des Objets Connectés.xlsx',
        iot_objects_array)

    # Create the security report.
    security_report = SecurityReport.SecurityReport()
    security_report.populate_member_values(iot_objects_array)
    writepptreport.write_dashboard_file(
        'Dashboard de Suivi de Sécurité des Objets Connectés.pptx',
        security_report)
        
    # Save the current SecurityReport to an xml file.
    security_report.save_to_file()


if __name__ == '__main__':
    main()
    print("Done")
