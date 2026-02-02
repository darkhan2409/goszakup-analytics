import sys
import os

def main():
    while True:
        print("\n" + "="*30)
        print("=== Goszakup Аналитика ===")
        print("="*30)
        print("1. Выгрузить все договоры (get_contracts)")
        print("2. Сгенерировать аналитический отчет (generate_report)")
        print("3. Показать сводку по объявлениям (get_announcements)")
        print("q. Выход")
        
        choice = input("\nВыберите действие: ").strip().lower()
        
        if choice == '1':
            print("\nЗапуск выгрузки договоров...")
            os.system("python src/get_contracts.py")
        elif choice == '2':
            print("\nЗапуск генерации отчета...")
            os.system("python src/generate_report.py")
        elif choice == '3':
            print("\nЗапуск сводки по объявлениям...")
            os.system("python src/get_announcements.py")
        elif choice == 'q':
            print("Выход из программы.")
            break
        else:
            print("Неверный выбор. Пожалуйста, попробуйте еще раз.")

if __name__ == "__main__":
    main()
