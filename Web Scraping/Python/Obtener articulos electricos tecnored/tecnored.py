#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# Scraper simple y lineal para:
# https://www.tiendatecnored.cl/materiales-electricos
# Extrae: page, sku, name, url, vat_info, price, price_clean.
# Fuerza decodificación UTF-8 y guarda resultados en un archivo Excel (.xlsx) usando pandas+openpyxl.
# Uso: python scrape_tecnored_to_excel.py

import requests
from bs4 import BeautifulSoup
from urllib.parse import urljoin
import time
import sys
import re

try:
    import pandas as pd
except Exception as e:
    print("Error: pandas no está instalado. Instálalo con: pip install pandas openpyxl", file=sys.stderr)
    raise

BASE_URL = 'https://www.tiendatecnored.cl/materiales-electricos'
HEADERS = {
    "User-Agent": "Mozilla/5.0 (compatible; simple-scraper/1.0)"
}
OUT_XLSX = 'tecnored_materiales.xlsx'
DELAY = 1.0

page = 1
rows = []

print("Iniciando scraping (decodificación forzada en UTF-8)...", file=sys.stderr)

while True:
    print(f"Obteniendo página {page}...", file=sys.stderr)
    try:
        r = requests.get(BASE_URL, params={'p': page}, headers=HEADERS, timeout=15)
        r.raise_for_status()
    except Exception as e:
        print(f"Error al pedir página {page}: {e}", file=sys.stderr)
        break

    try:
        html = r.content.decode('utf-8', errors='replace')
    except Exception:
        r.encoding = 'utf-8'
        html = r.text

    soup = BeautifulSoup(html, 'html.parser')
    product_anchors = soup.select('a.product-item-link')
    if not product_anchors:
        print("No se encontraron productos. Terminando.", file=sys.stderr)
        break

    for a in product_anchors:
        name = a.get_text(strip=True)
        href = a.get('href') or ''
        url = urljoin(BASE_URL, href)

        sku = ''
        parent = a
        for _ in range(6):
            parent = parent.parent
            if parent is None:
                break
            sku_div = parent.find('div', class_='product-sku-plp')
            if sku_div:
                span = sku_div.find('span')
                sku = span.get_text(strip=True) if span else sku_div.get_text(strip=True)
                break

        vat = ''
        if parent:
            vat_span = parent.find('span', class_='vat-info')
            if vat_span:
                vat = vat_span.get_text(strip=True)
        if not vat:
            gv = soup.find('span', class_='vat-info')
            if gv:
                vat = gv.get_text(strip=True)

        price = ''
        if parent:
            price_span = parent.find('span', class_='price')
            if price_span:
                price = price_span.get_text(strip=True)
        if not price:
            next_price = a.find_next('span', class_='price')
            if next_price:
                price = next_price.get_text(strip=True)

        price_clean = ''
        if price:
            cleaned = re.sub(r'[^\d,.-]', '', price)
            temp = cleaned
            if ',' in temp and '.' in temp:
                temp = temp.replace('.', '').replace(',', '.')
            else:
                temp = temp.replace('.', '')
                temp = temp.replace(',', '.')
            try:
                price_clean = float(temp)
            except Exception:
                price_clean = ''

        rows.append({
            'page': page,
            'sku': sku,
            'name': name,
            'url': url,
            'vat_info': vat,
            'price': price,
            'price_clean': price_clean
        })

    page += 1
    time.sleep(DELAY)

if rows:
    df = pd.DataFrame(rows, columns=['page', 'sku', 'name', 'url', 'vat_info', 'price', 'price_clean'])
    try:
        df.to_excel(OUT_XLSX, index=False, engine='openpyxl')
        print(f"Guardados {len(rows)} productos en {OUT_XLSX}", file=sys.stderr)
    except Exception as e:
        csv_fallback = 'tecnored_materiales_fallback.csv'
        df.to_csv(csv_fallback, index=False, encoding='utf-8-sig')
        print(f"No se pudo guardar .xlsx ({e}). Se guardó CSV en {csv_fallback}", file=sys.stderr)
else:
    print("No se extrajo ningún producto.", file=sys.stderr)