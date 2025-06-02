# CSV to XLS Converter for Product Data

This script converts a CSV file (commonly exported from e-commerce platforms like Shopify) to an XLS format with specific columns required by a product import template.

## Requirements

- Python 3.6 or higher
- pandas
- openpyxl

## Installation

1. Clone or download this repository
2. Install the required dependencies:

```bash
pip install -r requirements.txt
```

## Usage

### Basic usage

```bash
python csv_to_xls.py products_export_1.csv
```

This will create an output file with the same name but .xls extension (products_export_1.xls).

### Specify output file

```bash
python csv_to_xls.py products_export_1.csv output_template.xls
```

## Output Format

The script maps data from your CSV to the following columns:

1. Referência (código fornecedor) - Maps from "Handle"
2. Não importar os dados da coluna - Left empty
3. Código do produto (ID Tray) - Maps from "Variant SKU" or other ID fields
4. Nome do produto - Maps from "Title"
5. Preço de venda em reais - Maps from "Variant Price"
6. Preço de custo em reais - Maps from "Cost per item"
7. Estoque do produto - Maps from inventory fields
8. Exibir produto ativo - Maps from "Published" (converted to "S" for true and "N" for false)
9. NCM do produto - Looks for an NCM field
10. Código EAN/GTIN/UPC - Maps from "Variant Barcode"
11. Prazo de disponibilidade - Default is 5 days
12. SEO - Endereço do produto (URL) - Maps from "Handle" or generates from title
13. Peso do produto (gramas) - Maps from "Variant Grams"
14. Largura (cm) - Maps from dimensional metadata
15. Altura (cm) - Maps from dimensional metadata
16. Comprimento (cm) - Maps from dimensional metadata

## Notes

- The script has fallback options for many fields if the primary mapping isn't available
- Weights are automatically converted from kg to grams if needed
- Default values are provided when data is missing