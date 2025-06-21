from flask import Flask, request, jsonify, render_template
from flask_cors import CORS
from openpyxl import Workbook, load_workbook
import os
from datetime import datetime

app = Flask(__name__,
            template_folder='Dotaesmart', 
            static_folder='Dotaesmart') 

CORS(app)

CARPETA_PEDIDOS = "pedidos"

ARCHIVO_EXCEL = os.path.join(CARPETA_PEDIDOS, "pedidos.xlsx")

@app.route('/')
def index():
    try:
        return render_template('dotasmesmart.html') 
    except Exception as e:

        print(f"Error al renderizar la plantilla: {e}")
        return "<h1>¡La aplicación Dotaesmart está funcionando!</h1><p>Parece que hay un problema al cargar la página principal. Por favor, verifica el nombre del archivo HTML o la configuración de la carpeta 'templates'.</p>", 500


@app.route("/guardar_pedido", methods=["POST"])
def guardar_pedido():
    try:
        datos = request.json
        nombre = datos.get("nombre")
        direccion = datos.get("direccion")
        contacto = datos.get("contacto")
        carrito = datos.get("carrito", [])

   
        os.makedirs(os.path.join("Dotaesmart", CARPETA_PEDIDOS), exist_ok=True)

        ruta_excel_final = os.path.join("Dotaesmart", ARCHIVO_EXCEL)

        if not os.path.exists(ruta_excel_final):
            wb = Workbook()
            ws = wb.active
            ws.title = "Pedidos"
            ws.append(["Fecha", "Nombre", "Dirección", "Contacto", "Producto", "Cantidad", "Precio Unitario", "Subtotal"])
            wb.save(ruta_excel_final)

        wb = load_workbook(ruta_excel_final)
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

        wb.save(ruta_excel_final)
        return jsonify({"mensaje": "Pedido guardado con éxito"}), 200

    except Exception as e:
        print("Error al guardar el pedido:", e)
        return jsonify({"error": "Error interno del servidor", "detalle": str(e)}), 500

