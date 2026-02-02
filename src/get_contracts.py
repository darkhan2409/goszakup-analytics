import requests
import pandas as pd
import os
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment
from openpyxl.utils.dataframe import dataframe_to_rows
from config import TOKEN, BASE_URL, PAGE_LIMIT, BIN_COMPANY, FIN_YEAR, REPORTS_DIR

def get_contracts(bin_company, fin_year):
    """Получение договоров через GraphQL (заказчик, по финансовому году)"""
    
    headers = {
        "Authorization": f"Bearer {TOKEN}",
        "Content-Type": "application/json"
    }
    
    all_contracts = []
    after = 0
    
    while True:
        query = {
            "query": """
                query($limit: Int, $after: Int, $filter: ContractFiltersInput!) {
                    Contract(limit: $limit, after: $after, filter: $filter) {
                        id
                        contractNumber
                        signDate
                        contractSum
                        contractSumWnds
                        faktSum
                        supplierBiin
                        descriptionRu
                        finYear
                        Supplier {
                            nameRu
                        }
                        RefContractStatus {
                            nameRu
                        }
                        RefSubjectType {
                            nameRu
                        }
                        RefContractType {
                            nameRu
                        }
                        FaktTradeMethods {
                            nameRu
                        }
                        TrdBuy {
                            numberAnno
                        }
                        ContractUnits {
                            Plans {
                                amount
                            }
                        }
                    }
                }
            """,
            "variables": {
                "limit": PAGE_LIMIT,
                "after": after,
                "filter": {
                    "customerBin": bin_company,
                    "finYear": fin_year
                }
            }
        }
        
        response = requests.post(f"{BASE_URL}/v3/graphql", json=query, headers=headers)
        data = response.json()
        
        if "errors" in data:
            print(f"Ошибка API: {data['errors']}")
            break
        
        contracts = data.get("data", {}).get("Contract", [])
        if not contracts:
            break
            
        all_contracts.extend(contracts)
        
        page_info = data.get("extensions", {}).get("pageInfo", {})
        if not page_info.get("hasNextPage", False):
            break
        after = page_info.get("lastId", 0)
        
        print(f"Загружено: {len(all_contracts)} договоров...")
    
    return all_contracts

def format_number(value):
    """Форматирование числа с 2 знаками после запятой"""
    if value is None:
        return None
    try:
        return round(float(value), 2)
    except:
        return value

def get_plan_amount(contract):
    """Получение общей плановой суммы из предметов договора"""
    units = contract.get("ContractUnits", [])
    if not units:
        return None
    total = 0
    for unit in units:
        plans = unit.get("Plans")
        if plans and plans.get("amount"):
            total += plans.get("amount", 0)
    return total if total > 0 else None

def save_to_excel(contracts, filename):
    """Сохранение в Excel with форматированием"""
    
    # Порядок столбцов
    columns = [
        "№",
        "Номер договора в реестре договоров",
        "Номер закупки",
        "Описание",
        "Вид предмета",
        "Тип договора",
        "Статус",
        "Фактический способ закупки",
        "Финансовый год",
        "Общая плановая сумма договора",
        "Сумма без НДС",
        "Факт. сумма",
        "Наименование поставщика",
        "Дата заключения"
    ]
    
    rows = []
    for idx, c in enumerate(contracts, start=1):
        rows.append({
            "№": idx,
            "Номер договора в реестре договоров": c.get("contractNumber"),
            "Номер закупки": c.get("TrdBuy", {}).get("numberAnno") if c.get("TrdBuy") else None,
            "Описание": c.get("descriptionRu"),
            "Вид предмета": c.get("RefSubjectType", {}).get("nameRu") if c.get("RefSubjectType") else None,
            "Тип договора": c.get("RefContractType", {}).get("nameRu") if c.get("RefContractType") else None,
            "Статус": c.get("RefContractStatus", {}).get("nameRu") if c.get("RefContractStatus") else None,
            "Фактический способ закупки": c.get("FaktTradeMethods", {}).get("nameRu") if c.get("FaktTradeMethods") else None,
            "Финансовый год": c.get("finYear"),
            "Общая плановая сумма договора": format_number(get_plan_amount(c)),
            "Сумма без НДС": format_number(c.get("contractSum")),
            "Факт. сумма": format_number(c.get("faktSum")),
            "Наименование поставщика": c.get("Supplier", {}).get("nameRu") if c.get("Supplier") else None,
            "Дата заключения": c.get("signDate")[:10] if c.get("signDate") else None,
        })
    
    df = pd.DataFrame(rows, columns=columns)
    
    # Создаём книгу Excel
    wb = Workbook()
    ws = wb.active
    
    # Стили
    font_normal = Font(name='Times New Roman', size=12)
    font_bold = Font(name='Times New Roman', size=12, bold=True)
    alignment_center = Alignment(horizontal='center', vertical='center', wrap_text=True)
    alignment_left = Alignment(horizontal='left', vertical='center', wrap_text=True)
    
    # Индексы столбцов (1-based)
    description_col = 4  # "Описание"
    sum_cols = [10, 11, 12]  # "Общая плановая сумма договора", "Сумма без НДС", "Факт. сумма"
    
    # Записываем данные
    for r_idx, row in enumerate(dataframe_to_rows(df, index=False, header=True), 1):
        for c_idx, value in enumerate(row, 1):
            cell = ws.cell(row=r_idx, column=c_idx, value=value)
            
            # Выравнивание: Описание по левому краю (кроме заголовка)
            if c_idx == description_col and r_idx > 1:
                cell.alignment = alignment_left
            else:
                cell.alignment = alignment_center
            
            # Заголовки - жирный шрифт
            if r_idx == 1:
                cell.font = font_bold
            else:
                cell.font = font_normal
            
            # Числовой формат with разделителями разрядов
            if c_idx in sum_cols and r_idx > 1:
                cell.number_format = '#,##0.00'
    
    # Автоширина колонок
    for column in ws.columns:
        max_length = 0
        column_letter = column[0].column_letter
        for cell in column:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(str(cell.value))
            except:
                pass
        adjusted_width = min(max_length + 2, 50)
        ws.column_dimensions[column_letter].width = adjusted_width
    
    wb.save(filename)
    print(f"Сохранено: {filename}")

if __name__ == "__main__":
    print(f"Поиск договоров заказчика {BIN_COMPANY} за {FIN_YEAR} год...")
    
    contracts = get_contracts(BIN_COMPANY, FIN_YEAR)
    
    if contracts:
        filename = os.path.join(REPORTS_DIR, f"contracts_{BIN_COMPANY}_{FIN_YEAR}.xlsx")
        save_to_excel(contracts, filename)
        print(f"\nГотово! Найдено договоров: {len(contracts)}")
    else:
        print("Договоры не найдены.")
