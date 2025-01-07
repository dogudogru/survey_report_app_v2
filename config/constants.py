PARTY_MAPPING = {
    'Cumhuriyet Halk Partisi (CHP)': 'CHP',
    'Adalet ve Kalkınma Partisi (AKP)': 'AK Parti',
    'Yeşil Sol Parti (YSP)/ Halkların Demokratik Partisi (HDP)': 'DEM Parti',
    'Milliyetçi Hareket Partisi (MHP)': 'MHP',
    'Zafer Partisi': 'Zafer Partisi',
    'İYİ Parti': 'İYİ Parti',
    'Yeniden Refah Partisi': 'Yeniden Refah Partisi',
    'Anahtar Parti': 'Anahtar Parti',
    'Oy kullanmayacağım': 'Oy kullanmayacağım',
    'Kararsızım': 'Kararsızım'
}

PARTY_PAIRS = {
    16: ['CHP', 'AK Parti'],
    17: ['DEM Parti', 'MHP'],
    18: ['İYİ Parti', 'Kararsız']
}

# Education breakdown slides (19-22)
EDUCATION_CHART_MAPPING = {
    19: ['AK Parti', 'CHP'],
    20: ['DEM Parti', 'İYİ Parti'],
    21: ['MHP', 'Kararsız'],
    22: ['Oy kullanmayacağım']
}

# Age breakdown slides (22-25)
AGE_CHART_MAPPING = {
    22: ['AK Parti'],
    23: ['CHP', 'DEM Parti'],
    24: ['İYİ Parti', 'MHP'],
    25: ['Kararsız', 'Oy kullanmayacağım']
}

# 2023 party breakdown slide
PARTY_2023_SLIDE = 26

MAIN_PARTIES = set(PARTY_MAPPING.values())