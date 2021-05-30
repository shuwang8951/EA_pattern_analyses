class News:
    url = ''

    title = ''

    content = ''

    date = ''

    query_word = ''

    def __init__(self, url, title, date, content, query_word):
        self.url = url
        self.title = title
        self.date = date
        self.content = content
        self.query_word = query_word
