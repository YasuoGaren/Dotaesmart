from flask import Flask, request, jsonify
from flask_cors import CORS
from openpyxl import load_workbook
from datetime import datetime
import os

app = Flask(__name__)
CORS(app)

ARCHIVO_EXCEL = "pedidos/pedidos.xlsx"

@app.route('/guardar_pedido', methods=['POST'])
def guardar_pedido():
    try:
        data = request.get_json()

        nombre = data.get("nombre")
        direccion = data.get("direccion")
        contacto = data.get("contacto")
        carrito = data.get("carrito")

        if not all([nombre, direccion, contacto, carrito]):
            return jsonify({"error": "Faltan datos"}), 400

        wb = load_workbook(ARCHIVO_EXCEL)
        ws = wb.active

        for producto in carrito:
            ws.append([
                datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                nombre,
                direccion,
                contacto,
                producto['nombre'],
                producto['cantidad'],
                producto['precio'],
                producto['precio'] * producto['cantidad']
            ])

        wb.save(ARCHIVO_EXCEL)

        return jsonify({"mensaje": "Pedido guardado con éxito"})

    except Exception as e:
        print("Error al guardar el pedido:", str(e))
        return jsonify({"error": "Error interno del servidor"}), 500

if __name__ == '__main__':
    app.run(debug=True, port=5000)


