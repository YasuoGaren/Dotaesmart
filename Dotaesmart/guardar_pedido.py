from flask import Flask, request, jsonify
from flask_cors import CORS
from openpyxl import Workbook, load_workbook
import os
from datetime import datetime

app = Flask(__name__)
CORS(app)


CARPETA_PEDIDOS = "pedidos"
ARCHIVO_EXCEL = os.path.join(CARPETA_PEDIDOS, "pedidos.xlsx")

@app.route("/guardar_pedido", methods=["POST"])
def guardar_pedido():
    try:
        datos = request.json
        nombre = datos.get("nombre")
        direccion = datos.get("direccion")
        contacto = datos.get("contacto")
        carrito = datos.get("carrito", [])

        os.makedirs(CARPETA_PEDIDOS, exist_ok=True)


        if not os.path.exists(ARCHIVO_EXCEL):
            wb = Workbook()
            ws = wb.active
            ws.title = "Pedidos"
            ws.append(["Fecha", "Nombre", "Dirección", "Contacto", "Producto", "Cantidad", "Precio Unitario", "Subtotal"])
            wb.save(ARCHIVO_EXCEL)

        wb = load_workbook(ARCHIVO_EXCEL)
        ws = wb.active

        for item in carrito:
            nombre_producto = item["nombre"]
            cantidad = item["cantidad"]
            precio = item["precio"]
            subtotal = cantidad * precio
            ws.append([
                datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                nombre,
                direccion,
                contacto,
                nombre_producto,
                cantidad,
                precio,
                subtotal
            ])

        wb.save(ARCHIVO_EXCEL)
        return jsonify({"mensaje": "Pedido guardado con éxito"}), 200

    except Exception as e:
        print("Error al guardar el pedido:", e)
        return jsonify({"error": "Error interno del servidor"}), 500

if __name__ == "__main__":
    app.run(debug=True, port=5000)
