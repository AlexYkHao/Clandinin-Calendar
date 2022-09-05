from event_manager import EventManager

def main():
    TOKEN_PATH = 'creds/token.json'
    em = EventManager(TOKEN_PATH, 'clandinin0dev', 'test/test.xlsx')
    #em.remove_events_created_by(email='test@gmail.com')
    em.remove_events_id_with(id_prefix='clandinin0dev')


if __name__ == '__main__':
    main()