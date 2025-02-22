import pandas as pd
import win32com.client as win32

def export_outlook_contacts_to_excel(file_path):
    # Подключение к приложению Outlook
    outlook = win32.Dispatch('Outlook.Application')
    namespace = outlook.GetNamespace("MAPI")
    contacts_folder = namespace.GetDefaultFolder(10)  # 10 соответствует папке контактов в Outlook

    # Создание пустого DataFrame для хранения данных контактов
    contacts_data = {
        'Full Name': [],
        'Organization': [],
        'Job Title': [],
        'Email': []
    }

    # Итерация по всем контактам в папке
    for contact in contacts_folder.Items:
        if contact.Class == 40:  # 40 соответствует контактам в Outlook
            # Добавление данных контакта в DataFrame
            contacts_data['Full Name'].append(contact.FullName)
            contacts_data['Organization'].append(contact.CompanyName)
            contacts_data['Job Title'].append(contact.JobTitle)
            contacts_data['Email'].append(contact.Email1Address)

    # Создание DataFrame из собранных данных
    contacts_df = pd.DataFrame(contacts_data)

    # Экспорт DataFrame в файл Excel
    contacts_df.to_excel(file_path, index=False)
    print("Контакты успешно выгружены в Excel.")

if __name__ == "__main__":
    file_path = "data2.xlsx"
    export_outlook_contacts_to_excel(file_path)
