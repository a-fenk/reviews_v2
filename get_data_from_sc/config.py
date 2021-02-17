class Config:
    LIMIT_MASTERS = None

    MASTERS_SHEET = 'Masters_URL'
    REVIEWS_SHEET = 'All_reviews'
    SC_SHEET = 'SC'
    TAGS_SHEET = 'Шаблон тегов'

    TAG_WORD_COLUMN = 'Слова тегов'
    TAG_NAME_COLUMN = 'Name_tag'
    TAG_LEVEL_COLUMN = 'уровень'

    REVIEWS_COLUMNS = [
        'Отзыв',
        'Masters_URL',
        'ID section',
        'ID container',
        'Name section',
        '№ заказа',
        'Гео район',
        'Гео метро',
        'Corrected',
        'Кол-во отзывов',
        'Кол-во отзывов Corrected - TRUE',
    ]
    REVIEWS_SEARCH_RANGE = {
        'from': 'B',
        'to': 'B',
    }

    SC_COLUMNS = [
        'id container',
        'Address',
        'H1-1',
        'Id section 1'
    ]
    SC_SEARCH_RANGE = {
        'from': 'E',
        'to': 'N',
    }

    RESULT_COLUMNS = REVIEWS_COLUMNS + [
        'Address',
        'H1-1',
        'New ID container',
    ]

    RESULT_FILE_NAME = 'temp.xlsx'
    SOURCE_FILE_NAME = 'source.xlsx'

