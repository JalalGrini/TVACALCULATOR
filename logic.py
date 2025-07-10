def calculate_ht_tva(ttc, tva_rate):
    ht = round(ttc / (1 + tva_rate / 100), 2)
    tva = round(ttc - ht, 2)
    return ht, tva