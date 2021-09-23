# Tratamento-Dados
Repositório destinado a armazenar códigos que tem como objetivo manipular e tratar dados.

## ANALISE_BASE_CLIENTES.py

Utilizado para realizar carregar duas bases de dados e então mesclar ela para realizar uma análise rápida e então armazenar numa tabela no excel.

## ANÁLISE VISTORIA.py

Elaborado para analisar uma vistoria dos técnicos e assim realizar a medição de produtividade. A análise se baseia na hora em que cada item foi registrado na viatoria.

## CRUZAMENTO_POSTES.py

Tem como principal função identificar a sinergia entre duas bases de pontos geográficos, realizando calculo das distancias dos pontos. Possui função para cáculo em UTM ou em Coordenadas Decimais.

```python
#DECIMAL
def dist(lat1,lon1,lat2,lon2):
    try:
        lat1,lon1,lat2,lon2 = map(radians,[float(lat1),float(lon1),float(lat2),float(lon2)])
        R = 6371

        dlon = lon2 - lon1
        dlat = lat2 - lat1

        a = sin(dlat / 2)**2 + cos(lat1) * cos(lat2) * sin(dlon / 2)**2
        c = 2 * asin(sqrt(a))

        km = R * c
    except:
        km = 1000
        
    return km * 1000
```
ou
```python
#UTM
def distance_cartesian(x1, y1, x2, y2):
    dx = x1 - x2
    dy = y1 - y2

    return sqrt(dx * dx + dy * dy)
```

## MAPEAMENTO.py

Foi desenvolvido para realizar um mapeamento de rede FTTH através de uma planilha de Ponta A e Ponta B, indicando as conexões entre caixas de atendimento.

## License
[MIT](https://choosealicense.com/licenses/mit/)
