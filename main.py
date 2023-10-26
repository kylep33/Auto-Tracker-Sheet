import openpyxl
import Points_List_Reader
import install_sheet_creator


def main():

    # points list 1
    file_path = r"C:\Users\delta\PycharmProjects\Project Tracking Excel Sheet\testing_dir\Points List Template.xlsx"
    points_list_workbook = Points_List_Reader.load_excel_file(file_path)

    #get job name!
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
        exit(-1)

    # Extract unit type from the title
    unit_types = Points_List_Reader.extract_unit_type(title)
    print("Extracted Unit Types:", unit_types)

    # Create the IP-OP dictionary
    ip_op_dict = Points_List_Reader.create_ip_op_dict(points_list_workbook.active)

    # Set up workbook for the unit type
    target_directory = r'C:\Users\delta\PycharmProjects\Project Tracking Excel Sheet'
    workbook, full_path = install_sheet_creator.create_excel_workbook(job_name, unit_type, target_directory)

    # Build the sheet
    install_sheet_creator.build_workbook(workbook, full_path,job_name, unit_type, int(num_of_units), ip_op_dict)

    # Close the workbook
    install_sheet_creator.close_workbook(workbook)



if __name__ == "__main__":
    main()