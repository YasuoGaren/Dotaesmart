from flask import Flask, request, jsonify
from flask_cors import CORS
from openpyxl import Workbook, load_workbook
from datetime import datetime
import os

app = Flask(__name__)
CORS(app)  # Permitir peticiones desde archivos HTML locales

# Ruta local al archivo Excel (en tu carpeta sincronizada con Google Drive)
EXCEL_PATH = r'C:\Users\u1874e\Desktop\Dotaesmart\pedidos\pedidos.xlsx'

# Crear el archivo si no existe, con encabezados
if not os.path.exists(EXCEL_PATH):
    wb = Workbook()
    ws = wb.active
    ws.title = "Pedidos"
    ws.append(["Fecha", "Nombre", "Dirección", "Contacto", "Producto", "Cantidad", "Precio Unitario", "Total"])
    wb.save(EXCEL_PATH)

@app.route('/guardar_pedido', methods=['POST'])
def guardar_pedido():
    try:
        data = request.get_json()
        nombre = data.get('nombre')
        direccion = data.get('direccion')
        contacto = data.get('contacto')
        carrito = data.get('carrito', [])

        if not nombre or not direccion or not contacto or not carrito:
            return jsonify({"error": "Faltan datos"}), 400

        wb = load_workbook(EXCEL_PATH)
        ws = wb.active
        fecha = datetime.now().strftime('%Y-%m-%d %H:%M:%S')

        for item in carrito:
            total = item['cantidad'] * item['precio']
            ws.append([
                fecha, nombre, direccion, contacto,
                item['nombre'], item['cantidad'],
                item['precio'], total
            ])

        wb.save(EXCEL_PATH)
        return jsonify({"mensaje": "Pedido guardado exitosamente"}), 200

    except Exception as e:
        print(f"Error al guardar el pedido: {e}")
        return jsonify({"error": "Error interno"}), 500

if __name__ == '__main__':
    app.run(debug=True)

