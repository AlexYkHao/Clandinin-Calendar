from event_manager import EventManager

TOKEN_PATH = 'creds/token.json'
em = EventManager(TOKEN_PATH)
em.read_excel('test/test.xlsx')
em.setup_future_events()
