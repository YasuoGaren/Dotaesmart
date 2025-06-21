from flask import Flask, request, jsonify
from flask_cors import CORS
import os
from openpyxl import Workbook, load_workbook
from datetime import datetime

app = Flask(__name__)
CORS(app)

CARPETA_PEDIDOS = "pedidos"
ARCHIVO_EXCEL = os.path.join(CARPETA_PEDIDOS, "pedidos.xlsx")

os.makedirs(CARPETA_PEDIDOS, exist_ok=True)

@app.route('/guardar_pedido', methods=['POST'])
def guardar_pedido():
    try:
        datos = request.json
        nombre = datos['nombre']
        direccion = datos['direccion']
        contacto = datos['contacto']
        carrito = datos['carrito']


        if not os.path.exists(ARCHIVO_EXCEL):
            wb = Workbook()
            ws = wb.active
            ws.append(["Fecha", "Nombre", "Dirección", "Contacto", "Producto", "Cantidad", "Precio"])
        else:
            wb = load_workbook(ARCHIVO_EXCEL)
            ws = wb.active

        fecha = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        for item in carrito:
            ws.append([
                fecha,
                nombre,
                direccion,
                contacto,
                item['nombre'],
                item['cantidad'],
                item['precio']
            ])

        wb.save(ARCHIVO_EXCEL)

        return jsonify({"mensaje": "Pedido guardado correctamente"}), 200
    except Exception as e:
        print("Error al guardar el pedido:", e)
        return jsonify({"error": "Error al guardar el pedido"}), 500

if __name__ == '__main__':
    port = int(os.environ.get("PORT", 5000))
    app.run(host='0.0.0.0', port=port)


