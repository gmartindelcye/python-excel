"""
Reads an xlsx file, reads the A2 cell, saves the file in a folder structure year/month adding the date to
the file name.
If the month folder does not exist, it is created.
If the year folder does not exist, it is created.
"""
import os
import openpyxl


def save_file_by_date(file:str, basedir: str):
    try:
        # Get the current working directory
        cwd = os.getcwd()

        # Check if the folder structure exists, if not, create it
        if not os.path.exists(cwd + '/data'):
            os.makedirs(cwd + '/data')
            print('Folder "data" created')

        # Get the excel file name
        xfile = 'python_test.xlsx'

        # Open the excel file
        wb = openpyxl.load_workbook(xfile)

        # Get the sheet name
        sheet = wb.active

        # Get the cell value
        a2 = sheet['A2'].value

        # Get the year
        year = a2.year

        # Get the month
        month = a2.month
        mnth = str(month).zfill(2)

        # Get the day
        day = a2.day
        dy = str(day).zfill(2)

        # check if the year folder exists, if not, create it
        if not os.path.exists(cwd + '/data/' + str(year)):
            os.makedirs(cwd + '/data/' + str(year))
            print('Folder "' + str(year) + '" created')

        # check if the month folder exists, if not, create it
        if not os.path.exists(cwd + '/data/' + str(year) + '/' + mnth):
            os.makedirs(cwd + '/data/' + str(year) + '/' + mnth)
            print('Folder "' + mnth + '" created')

        # Save the file
        wb.save(cwd + '/data/' + str(year) + '/' + mnth + '/' + dy + '.xlsx')

        # Close the file
        wb.close()

    except:
        print('Error saving the file')
        return 1

    # Return 0 if everything is ok
    print('File saved')
    return 0


if __name__ == '__main__':
    save_file_by_date('python_test.xlsx', 'data')

