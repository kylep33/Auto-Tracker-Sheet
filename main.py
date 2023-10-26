import openpyxl
import Points_List_Reader
import install_sheet_creator

def read_points_list(file_path):
    # Load the points list workbook
    points_list_workbook = Points_List_Reader.load_excel_file(file_path)

    # Get job name
    job_name = Points_List_Reader.read_job(points_list_workbook.active)

    # Read and display title
    title = Points_List_Reader.read_title(points_list_workbook.active)
    print("Title:", title)

    title_info = Points_List_Reader.parse_title_text(title)
    if title_info:
        controller_type = title_info['controller_type']
        unit_type = title_info['unit_type']
        num_of_units = title_info['num_of_units']
        print("Controller Type:", controller_type)
        print("Unit Type:", unit_type)
        print("Number of Units:", num_of_units)
    else:
        print("No match found.")
        return None

    # Extract unit type from the title
    unit_types = Points_List_Reader.extract_unit_type(title)
    print("Extracted Unit Types:", unit_types)

    # Create the IP-OP dictionary
    ip_op_dict = Points_List_Reader.create_ip_op_dict(points_list_workbook.active)

    return job_name, unit_type, num_of_units, ip_op_dict

def create_install_sheets(job_name, unit_type, num_of_units, ip_op_dict):
    # Set up the target directory
    target_directory = r'C:\Users\delta\PycharmProjects\Project Tracking Excel Sheet'

    # Create the Excel workbook for the unit type
    workbook, full_path = install_sheet_creator.create_excel_workbook(unit_type, unit_type, target_directory)
    install_sheet_creator.create_unit_sheets(workbook, unit_type)

    # Build the sheet
    install_sheet_creator.build_workbook(workbook, full_path, job_name, unit_type, int(num_of_units), ip_op_dict)

    # Close the workbook
    install_sheet_creator.close_workbook(workbook)

def main():
    # List of file paths for points lists
    file_paths = [
        r"C:\Users\delta\PycharmProjects\Project Tracking Excel Sheet\testing_dir\Points List Template1.xlsx",
        r"C:\Users\delta\PycharmProjects\Project Tracking Excel Sheet\testing_dir\Points List Template2.xlsx",
        # Add more file paths as needed
    ]

    for file_path in file_paths:
        result = read_points_list(file_path)
        if result:
            job_name, unit_type, num_of_units, ip_op_dict = result
            create_install_sheets(job_name, unit_type, num_of_units, ip_op_dict)


if __name__ == "__main__":
    main()
