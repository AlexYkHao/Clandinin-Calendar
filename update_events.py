from event_manager import EventManager

def main():
    TOKEN_PATH = 'creds/token.json'
    em = EventManager(TOKEN_PATH, 'clandinin0dev', 'test/test.xlsx')
    col_mapper = {
        'ID': 'A',
        'Need Update': 'J',
        'Error State': 'K'
    }
    em.update_from_excel(col_mapper, allow_overlap=False)


if __name__ == '__main__':
    main()
