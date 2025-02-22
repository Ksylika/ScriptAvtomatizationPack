import pandas as pd
import win32com.client as win32

def create_outlook_contacts_from_excel(file_path):
    # Загрузка данных из файла Excel
    df = pd.read_excel(file_path)

    # Подключение к приложению Outlook
    outlook = win32.Dispatch('Outlook.Application')
    namespace = outlook.GetNamespace("MAPI")
    contacts_folder = namespace.GetDefaultFolder(10)  # 10 соответствует папке контактов в Outlook

    # Создание контактов из данных Excel
    for index, row in df.iterrows():
        full_name = row['Full Name']
        organization = row['Organization']
        job_title = row['Job Title']
        email = row['Email']

        # Создание нового контакта
        contact = outlook.CreateItem(2)  # 2 соответствует типу контакта в Outlook
        print(full_name)
        # Заполнение данных контакта
        contact.FullName = full_name
        contact.CompanyName = organization
        contact.JobTitle = job_title
        contact.Email1Address = email

        # Сохранение контакта
        contact.Save()
        print(f"Контакт '{full_name}' успешно создан в Outlook.")

    print("Все контакты успешно созданы.")

if __name__ == "__main__":
    file_path = "data.xlsx"
    create_outlook_contacts_from_excel(file_path)
