"""
File containing a class with basic
methods for simulating working
with files to simplify testing
"""
import os
import docx
import openpyxl


class CommonMethodsTesting:
    @staticmethod
    def creare_word_and_excel_file() -> str:
        word_path = CommonMethodsTesting.create_word_file()
        excel_path = CommonMethodsTesting.create_excel_file()
        return word_path, excel_path

    @staticmethod
    def create_word_file() -> str:
        """
        Method for creating an empty Word
        file 'My_docx1.docx' and returning its path
        Returns:
            str: the path to the file
        """
        file_path = os.path.join(os.path.join(os.path.expanduser('~')), 'Desktop')
        file_path += '/My_docx1.docx'
        test_docx = docx.Document()
        test_docx.save(file_path)
        file_path = '/'.join(file_path.split('\\'))
        return file_path

    @staticmethod
    def create_excel_file() -> str:
        """
        Method for creating an empty WExcel
        file 'My_xlsx1.xlsx' and returning its path
        Returns:
            str: the path to the file
        """
        file_path = os.path.join(os.path.join(os.path.expanduser('~')), 'Desktop')
        file_path += '/My_xlsx1.xlsx'
        test_xlsx = openpyxl.Workbook()
        test_xlsx.create_sheet('Лист1')
        test_xlsx.save(file_path)
        file_path = '/'.join(file_path.split('\\'))
        return file_path

    @staticmethod
    def remove_word_and_excel_file(
        word_path: str,
        excel_path: str
    ) -> None:
        CommonMethodsTesting.remove_file(word_path)
        CommonMethodsTesting.remove_file(excel_path)

    @staticmethod
    def remove_file(file_path: str) -> None:
        """
        Method deleting a Word
        file used for testing
        Args:
            path (str): the path to the file
        """
        os.remove(file_path)
