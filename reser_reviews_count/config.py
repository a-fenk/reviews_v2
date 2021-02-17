class Config:
    SOURCE_FILE = 'temp.xlsx'
    RESULT_FILE = 'result.xlsx'
    REVIEWS_SHEET = 'All_reviews'  # лист в котором ведём работу
    GEO_SHEET = 'ГЕО'  # Лист с гео данными

    GEO_DISTRICT_COLUMN = 'Гео район'
    GEO_METRO_COLUMN = 'Гео метро'
    REVIEWS_COUNT_COLUMN = 'Кол-во отзывов'
    MASTER_COLUMN = 'Masters_URL'
    ID_SECTION_COLUMN = 'ID section'
    NEW_ID_CONTAINER_COLUMN = 'New ID container'
    ID_CONTAINER_COLUMN = 'ID container'
    H1_COLUMN = 'H1-1'

    LTE = 1  # less than OR equal
    GTE = 3  # greater than OR equal
    TARGET = 2
