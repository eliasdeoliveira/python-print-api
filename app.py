import traceback
from flask import Flask, request, jsonify
import win32print
import win32ui
import win32api
import win32con
import json
import textwrap  # Para quebra de linha automática
from datetime import datetime
from flask_cors import CORS

app = Flask(__name__)


def format_json_to_table(data):
    lines = []
    try:
        # Verifica se 'header' está presente no JSON
        divider = '=' * 39
        divider2 = '_' * 39
        lines.append(datetime.now().strftime("%d/%m/%Y %H:%M:%S"))
        lines.append(divider2)
        if 'header' in data:
            header = data['header']
            lines.append(header.get('companyName', 'Nome da Empresa'))
            # Adiciona o divisor e a quebra de linha juntos
            lines.append(divider)
            lines.append('CUPOM NÃO FISCAL'.center(40))
            if (header.get('numberOrder')):
                # Adiciona o divisor e a quebra de linha juntos
                lines.append(divider)
                lines.append(f"Número do Pedido: {
                             header.get('numberOrder')}".capitalize())
        else:
            # lines.append('Nome da Empresa')
            lines.append('CUPOM NÃO FISCAL'.center(50))
        lines.append(divider)  # Adiciona o divisor e a quebra de linha juntos

        # Adiciona as colunas
        # lines.append(f"{'Index':<5} {'Produto':<20} {'Cor':<10} {'Tamanho':<10} {'Qtd':<5} {'Preço':<10}")

        # Comanda
        if (data['command']):
            lines.append(f"COMANDA: {data['command']}")
            lines.append(divider)

        # Verifica se 'items' está presente e é uma lista
        if 'items' in data and isinstance(data['items'], list):
            # Ajuste '40' para a largura da sua impressora
            lines.append("PRODUTOS".center(50))
            lines.append(divider)
            count = 0
            for item in data['items']:
                # lines.append(f"# {item.get('index', ''):<5} {item.get('productName', ''):<20} {item.get('productColor', ''):<10} {
                #  item.get('productSize', ''):<10} {item.get('quantity', ''):<5} {item.get('totalPrice', ''):<10}")
                lines.append(f"# {item.get('index', '')}) {
                             item.get('productName', '')}")
                if (item.get('productColor')):
                    lines.append(f"Cor: {item.get('productColor', ''):<10} ")
                if (item.get('productSize')):
                    lines.append(
                        f"Tamanho: {item.get('productSize', ''):<10} ")
                if (item.get('unitPrice')):
                    lines.append(f"Preço Unitário: {
                                 item.get('unitPrice', ''):<10} ")
                if (item.get('quantity')):
                    lines.append(f"Quantidade: {
                                 item.get('quantity', ''):<10} ")
                if (item.get('totalPrice')):
                    lines.append(
                        f"Total: R$: {item.get('totalPrice', ''):<10} ")
                if (item.get('observations')):
                    lines.append(f"Observações: {
                                 item.get('observations', '')}")
                # Itens adicionais
                if 'additionalItems' in item and isinstance(item['additionalItems'], list) and len(item['additionalItems']) > 0:
                    # Pula duas linhas
                    lines.append("\n.")  # Linha em branco
                    lines.append("ADICIONAIS")
                    for additional in item['additionalItems']:
                        lines.append(f"# {additional.get('index', '')}) {
                            additional.get('productName', '')}")
                        if (additional.get('productColor')):
                            lines.append(
                                f"Cor: {additional.get('productColor', ''):<10} ")
                        if (additional.get('productSize')):
                            lines.append(
                                f"Tamanho: {additional.get('productSize', ''):<10} ")
                        if (additional.get('unitPrice')):
                            lines.append(f"Preço Unitário: {
                                         additional.get('unitPrice', ''):<10} ")
                        if (additional.get('quantity')):
                            lines.append(f"Quantidade: {
                                additional.get('quantity', ''):<10} ")
                        if (additional.get('totalPrice')):
                            lines.append(
                                f"Total: R$: {additional.get('totalPrice', ''):<10} ")
                        if (additional.get('observations')):
                            lines.append(
                                f"Observações: {additional.get('observations', ''):<10} ")
                count += 1
                if (count < len(data['items'])):
                    lines.append(divider2 + "\n.")  # Linha em branco
        # Verifica se 'paymentDetails' está presente
        if 'paymentDetails' in data:
            lines.append(divider)
            lines.append("DETALHES DE PAGAMENTO".center(30))
            lines.append(divider)
            payment = data['paymentDetails']
            if (payment.get('paymentForm')):
                lines.append(f"Forma de Pagamento: {
                    payment.get('paymentForm', 'Não especificado')}")
                lines.append(divider)
            if (payment.get('paymentMethod')):
                lines.append(f"Método de Pagamento: {payment.get(
                    'paymentMethod', 'Não especificado')}")
                lines.append(divider)
            if (payment.get('total')):
                lines.append(f"Total: {payment.get(
                    'total', 'Não especificado')}")
                lines.append(divider)
            if (payment.get('cashDepositAmount')):
                lines.append(f"Depósito em Dinheiro: {payment.get(
                    'cashDepositAmount', 'Não especificado')}")
                lines.append(divider)
            if (payment.get('cardNumber')):
                lines.append(f"Número de Parcelas: {payment.get(
                    'numberInstallments', 'Não especificado')}")
                lines.append(divider)
            if (payment.get('discount')):
                lines.append(f"Desconto: {payment.get(
                    'discount', 'Não especificado')}")
                lines.append(divider)
            if (payment.get('finalAmount')):
                lines.append(f"Valor Final: {payment.get(
                    'finalAmount', 'Não especificado')}")
                lines.append(divider)
            if (payment.get('change')):
                lines.append(f"Troco: {payment.get(
                    'change', 'Não especificado')}")

        # Finaliza a impressão
        lines.append("AGRADECEMOS A PREFERÊNCIA.".center(30))
        lines.append("\x1D\x56\x00")  # Eject paper
        # lines.append("\x1D\x56\x00")  # Eject paper

    except Exception as e:
        print(f"Error in format_json_to_table: {e}")
        print(traceback.format_exc())
        raise

    return "\n".join(lines)


def print_text(text, printer_name, font_size, margins, type):
    try:
        # Abre a impressora especificada
        hprinter = win32print.OpenPrinter(printer_name)

        # Cria o dispositivo de contexto (DC) para a impressora
        hdc = win32ui.CreateDC()
        hdc.CreatePrinterDC(printer_name)

        # Configura a fonte
        font = win32ui.CreateFont({
            "name": "Arial",
            "height": font_size,
            "weight": win32con.FW_NORMAL
        })
        hdc.SelectObject(font)
        hdc.SetTextColor(win32api.RGB(0, 0, 0))  # Preto
        hdc.SetBkMode(win32con.TRANSPARENT)

        # Obtém a resolução da impressora (DPI)
        dpi_x = hdc.GetDeviceCaps(win32con.LOGPIXELSX)
        dpi_y = hdc.GetDeviceCaps(win32con.LOGPIXELSY)

        # Converte a largura de 50mm para pixels (50mm * DPI / 25.4mm por polegada)
        max_width_mm = 50
        max_width_px = int(max_width_mm * dpi_x / 25.4)

        # Aplica as margens
        margin_left = max(margins.get('left', 10), 10)
        margin_top = max(margins.get('top', 10), 10)
        text_x = margin_left
        text_y = margin_top

        # Quebra o texto em linhas de acordo com a largura máxima
        text_lines = []
        for line in text.splitlines():
            # Ajusta o valor para o tamanho da fonte
            text_lines.extend(textwrap.wrap(line, width=max_width_px // 10))

        # Inicia o trabalho de impressão
        hdc.StartDoc("Print Job")
        hdc.StartPage()

        # Imprime o texto com margens e quebra automática de linha
        for line in text_lines:
            hdc.TextOut(text_x, text_y, line)
            text_y += font_size  # Ajuste o espaçamento entre linhas conforme necessário

        # Finaliza a página e o trabalho de impressão
        hdc.EndPage()
        hdc.EndDoc()

        # Fecha a impressora
        win32print.ClosePrinter(hprinter)
    except Exception as e:
        print(f"Error in print_text: {e}")
        print(traceback.format_exc())
        raise


def string_to_json(json_string):
    try:
        # Converte a string JSON para um objeto Python (dicionário)
        data = json.loads(json_string)
        return data
    except json.JSONDecodeError as e:
        # Captura erros de decodificação JSON
        print(f"Error decoding JSON: {e}")
        return None
    except Exception as e:
        # Captura qualquer outro erro
        print(f"Unexpected error: {e}")
        return None


# CORS(app, resources={r"/*": {"origins": "*"}}, methods=["GET", "POST"])
CORS(app, resources={r"/*": {"origins": "*"}})


@app.route('/printer', methods=['POST'])
def printRouter():
    print('Print request received')
    body = request.json
    print('### body: ', body)
    data = body['data']
    if not data:
        return jsonify({'error': 'No data provided'}), 400

    try:
        printer_name = data.get('printer_name', win32print.GetDefaultPrinter())
        font_size = data.get('font_size', 12)
        margins = data.get('margins', {})
        if (data.get('type') == 'cupom'):
            text = format_json_to_table(data['text'])
            print_text(text, printer_name, font_size, margins, data['type'])
        else:
            lines = []
            lines.append(data['text'])
            print_text(lines, printer_name, font_size, margins, data['type'])
        return jsonify({'status': 'success'}), 200
    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/status', methods=['GET'])
def status():
    return jsonify({'status': 'API is running'}), 200


if __name__ == '__main__':
    # app.run(debug=True)
    app.run(host='0.0.0.0', port=5000, debug=True)
