import subprocess
import os
import pandas as pd
import argparse

# Константы алфавитов
ENG = "abcdefghijklmnopqrstuvwxyz"
RUS = "абвгдежзийклмнопрстуфхцчшщъыьэюя"

# Алгоритм цезаря
def caesar_shift(text, shift):
    res = []
    for char in text:
        low_char = char.lower()
        if low_char in ENG:
            alphabet = ENG
            idx = alphabet.index(low_char)
            new_char = alphabet[(idx - shift) % len(alphabet)]
            res.append(new_char if char.islower() else new_char.upper())
        elif low_char in RUS:
            alphabet = RUS
            idx = alphabet.index(low_char)
            new_char = alphabet[(idx - shift) % len(alphabet)]
            res.append(new_char if char.islower() else new_char.upper())
        else:
            res.append(char)
    return "".join(res)

# Лингвистический анализ
def get_best_shift(text_fragment):
    markers = ['ул', 'д.', 'кв', 'пр', 'обл', 'г.', 'ш.']
    best_shift = 0
    max_matches = -1
    for s in range(33):
        decrypted = caesar_shift(text_fragment.lower(), s)
        matches = sum(1 for m in markers if m in decrypted)
        if matches > max_matches:
            max_matches = matches
            best_shift = s
    return best_shift

# Чтение из файлаы
def read_data_from_excel(file_path):
    # Читаем исходный файл (пропуская пустые строку 1 и колонку A)
    df = pd.read_excel(file_path, skiprows=1, usecols="B:D")
    return df

#Брутфорс хешей
def run_hashcat_simple_sha1(hash_list, prefix="89"):
    hash_file = "hashes_to_crack.txt"
    with open(hash_file, "w") as f:
        for h in set(hash_list):
            if h: f.write(f"{h}\n")

    mask = f"{prefix}" + "?d" * (11 - len(prefix))
    command = [
        "hashcat", "-m", "100", "-a", "3",
        hash_file, mask, "--quiet", "--potfile-disable", "--force"
    ]

    try:
        result = subprocess.run(command, capture_output=True, text=True)
        cracked_data = {}
        for line in result.stdout.splitlines():
            if ":" in line:
                h_val, phone_val = line.split(":")
                cracked_data[h_val] = phone_val
        return cracked_data
    except Exception as e:
        print(f"Ошибка Hashcat: {e}")
        return {}
    finally:
        if os.path.exists(hash_file):
            os.remove(hash_file)

def process_main(file_path):
    df = read_data_from_excel(file_path)
    final_data = []
    hashes_only = []

    for _, row in df.iterrows():
        sha1_raw = str(row['Телефон']).strip()
        email_enc = str(row['email']).strip()
        address_enc = str(row['Адрес']).strip()

        shift = get_best_shift(address_enc)
        
        email_dec = caesar_shift(email_enc, shift)
        address_dec = caesar_shift(address_enc, shift)

        # Собираем данные в нужном порядке
        final_data.append({
            'Телефон': "...", 
            'email': email_dec,
            'Адрес': address_dec,
            'Ключ': shift,
            '_hash_key': sha1_raw
        })
        hashes_only.append(sha1_raw)

    cracked_map = run_hashcat_simple_sha1(hashes_only)

    for rec in final_data:
        rec['Телефон'] = cracked_map.get(rec.pop('_hash_key'), "Не найден")

    return final_data

# Сохраняем данные в файл
def write_data_to_excel(output_path, processed_records):
    df_output = pd.DataFrame(processed_records)
    
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        df_output.to_excel(writer, index=False)
        
        # Получаем объект листа для применения AutoFit
        worksheet = writer.sheets['Sheet1']
        
        for col in worksheet.columns:
            max_length = 0
            column = col[0].column_letter # Буква колонки
            for cell in col:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass
            adjusted_width = (max_length + 2)
            worksheet.column_dimensions[column].width = adjusted_width

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Деобфускация данных из Excel с помощью Hashcat.")
    parser.add_argument("-i", "--input", required=True, help="Путь к входному Excel файлу")
    parser.add_argument("-o", "--output", default="output_decrypted.xlsx", help="Путь к выходному Excel файлу")
    
    args = parser.parse_args()

    if os.path.exists(args.input):
        print(f"Обработка файла: {args.input}")
        results = process_main(args.input)
        write_data_to_excel(args.output, results)
        print(f"Готово! Результаты сохранены в: {args.output}")
    else:
        print(f"Ошибка: Файл {args.input} не найден.")