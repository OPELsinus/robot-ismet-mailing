


















emails = {'ТОВАРИЩЕСТВО С ОГРАНИЧЕННОЙ ОТВЕТСТВЕННОСТЬЮ "ALMA TRADE DISTRIBUTION"': 'fortisline.elnar@mail.ru, uchet.fortis@mail.ru, akty.almatrade@gmail.com', 'ТОВАРИЩЕСТВО С ОГРАНИЧЕННОЙ ОТВЕТСТВЕННОСТЬЮ "FORTIS SKO"': 'rogacheva.1981@mail.ru'}

for key, emails_ in emails.items():
    print('----------')
    print(key)
    emls = []
    print(f"to={[email for email in emails_.split(',')]}, subject=f'Исмет Рассылка Тест', username=smtp_author)")

