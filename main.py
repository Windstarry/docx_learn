from docx import Document
from utils.governmentaffairs import GovernmentAffair,MaterialList
from utils.filefunctions import file_save_name,add_contents,down_material_lists



def main(startline,endline):
    for i in range(startline,endline):
        file_model = Document()
        governmentaffair = GovernmentAffair(i)
        print(governmentaffair.dept_name,governmentaffair.service_name)
        add_contents(file_model,governmentaffair)
        file_save_name(file_model,governmentaffair)
        down_material_lists(governmentaffair)


if __name__ == '__main__':
    startline = 1785
    endline = 1791
    main(startline,endline)

    
