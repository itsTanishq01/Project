from ImportToSQL import get_excel_file_path, import_data_to_mysql
from Report import export_data_to_excel as export_report1
from Report2 import export_data_to_excel as export_report2
from ReportUI import choose_report

def main():
    # Import data into MySQL
    excel_file_path_import = get_excel_file_path()

    if excel_file_path_import:
        import_data_to_mysql(excel_file_path_import)
        print("Data imported to MySQL.")
    else:
        print("No file selected for import. Exiting.")
        return

    # Choose which report to export using the UI
    def on_report_selected(report_number):
        if report_number == 1:
            export_report1()
        elif report_number == 2:
            export_report2()
        else:
            print("Invalid report number. Exiting.")

    choose_report(on_report_selected)

if __name__ == "__main__":
    main()
