from event_manager import EventManager

def main():
    TOKEN_PATH = 'creds/token.json'
    em = EventManager(TOKEN_PATH, 'clandinin0dev', 'test/test.xlsx')
    em.update_from_excel(allow_overlap=False)


if __name__ == '__main__':
    main()
