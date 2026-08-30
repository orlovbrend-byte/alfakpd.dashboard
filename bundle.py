# -*- coding: utf-8 -*-
"""Собирает однофайловый index.html: без CDN, открывается с диска."""
import json, io

def rd(p):
    return io.open(p, encoding='utf-8').read()

tpl = rd('index.template.html')
data = rd('../data.json')
# Данные внутри <script type=application/json> — единственная опасность
# это последовательность </script  внутри строки; экранируем её.
data = data.replace('</', '<\\/')

out = (tpl
  .replace('/*__CSS__*/', rd('style.css'))
  .replace('/*__CHARTJS__*/', rd('chart.umd.js'))
  .replace('/*__CORE__*/', rd('core.js'))
  .replace('/*__APP__*/', rd('app.js'))
  .replace('/*__DATA__*/', data))

io.open('index.html', 'w', encoding='utf-8').write(out)
print('index.html', len(out), 'байт')
