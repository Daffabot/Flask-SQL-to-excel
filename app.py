from flask import Flask, request, send_from_directory, jsonify, Response
import pymysql
import xlsxwriter
import os
import uuid
import time
from flasgger import Swagger, LazyString

app = Flask(__name__)

swagger_config = {
    "headers": [],
    "specs": [
        {
            "endpoint": "apispec_1",
            "route": "/apispec_1.json",
            "rule_filter": lambda rule: True,
            "model_filter": lambda tag: True,
        }
    ],
    "static_url_path": None,
    "swagger_ui": False,
    "specs_route": None
}
swagger = Swagger(app, config=swagger_config)

# --- Path untuk export file ---
EXPORT_FOLDER = os.path.join(app.root_path, 'static', 'exports')
os.makedirs(EXPORT_FOLDER, exist_ok=True)

# --- Konfigurasi DB ---
DB_CONFIG = {
    "host": os.environ.get("DB_HOST", "localhost"),
    "user": os.environ.get("DB_USER", "daffabot"),
    "password": os.environ.get("DB_PASSWORD", "261291"),
    "database": os.environ.get("DB_NAME", "web")
}

@app.route('/docs')
def custom_docs():
    html = """
    <!DOCTYPE html>
    <html lang="en">
    <head>
      <meta charset="UTF-8">
      <title>Daffabot API Docs</title>
      <link rel="stylesheet" href="https://unpkg.com/swagger-ui-dist/swagger-ui.css">
      <style>
        body { margin:0; background:#fcfcfc; }
        #swagger-ui .topbar { background:#ff6a00; }
      </style>
    </head>
    <body>
      <div id="swagger-ui"></div>
      <script src="https://unpkg.com/swagger-ui-dist/swagger-ui-bundle.js"></script>
      <script src="https://unpkg.com/swagger-ui-dist/swagger-ui-standalone-preset.js"></script>
      <script>
        window.onload = () => {
          SwaggerUIBundle({
            url: "/api/to-excel/apispec_1.json",
            dom_id: '#swagger-ui',
            presets: [
              SwaggerUIBundle.presets.apis,
              SwaggerUIStandalonePreset
            ],
            layout: "BaseLayout"
          });
        };
      </script>
    </body>
    </html>
    """
    return Response(html, mimetype="text/html")

# --- API Export Excel ---
@app.route('/export', methods=['POST'])
def export_excel():
    """
    Export hasil query MySQL ke file Excel
    ---
    tags:
      - Export
    parameters:
      - name: body
        in: body
        required: true
        schema:
          type: object
          properties:
            db_config:
              type: object
              properties:
                host:
                  type: string
                  example: "localhost"
                user:
                  type: string
                  example: "root"
                password:
                  type: string
                  example: "mypassword"
                database:
                  type: string
                  example: "test"
            query:
              type: string
              example: "SELECT * FROM users LIMIT 100"
            header:
              type: array
              items:
                type: string
              example: ["ID", "Nama", "Email"]
    responses:
      200:
        description: URL download Excel
        schema:
          type: object
          properties:
            download_url:
              type: string
            time:
              type: string
      400:
        description: Query salah / body tidak valid
      500:
        description: Error server
    """
    data = request.get_json()

    if not data or 'query' not in data:
        return jsonify({"error": "Missing 'query' in request body"}), 400

    query = data['query']
    custom_header = data.get('header', [])

    # Gunakan DB_CONFIG bawaan, bisa di override lewat request body
    db_config = DB_CONFIG.copy()
    if 'db_config' in data:
        db_config.update({k: v for k, v in data['db_config'].items() if v})

    try:
        connection = pymysql.connect(**db_config)
        cursor = connection.cursor()
        cursor.execute(query)
        rows = cursor.fetchall()
        column_names = [desc[0] for desc in cursor.description]

        if not rows:
            return jsonify({"info": "Query berhasil, tapi hasilnya kosong"}), 200

        final_header = [
            custom_header[i] if i < len(custom_header) else col
            for i, col in enumerate(column_names)
        ]

        start = time.time()
        filename = f"{uuid.uuid4().hex}.xlsx"
        filepath = os.path.join(EXPORT_FOLDER, filename)

        workbook = xlsxwriter.Workbook(filepath)
        worksheet = workbook.add_worksheet("Sheet1")

        for col_num, header in enumerate(final_header):
            worksheet.write(0, col_num, header)

        for row_num, row in enumerate(rows, start=1):
            for col_num, value in enumerate(row):
                worksheet.write(row_num, col_num, value)

        workbook.close()
        cursor.close()
        connection.close()
        end = time.time()

        return jsonify({
            "download_url": f"{request.host_url}download/{filename}",
            "time": f"Selesai dalam {end - start:.2f} detik"
        })

    except pymysql.err.ProgrammingError as e:
        return jsonify({"error": f"SQL query error: {str(e)}"}), 400

    except Exception as e:
        return jsonify({"error": str(e)}), 500

# --- Download file ---
@app.route('/download/<filename>')
def download_file(filename):
    """
    Download file Excel hasil export
    ---
    tags:
      - Export
    parameters:
      - name: filename
        in: path
        type: string
        required: true
        description: Nama file hasil export
    responses:
      200:
        description: File berhasil diunduh
        schema:
          type: file
    """
    return send_from_directory(EXPORT_FOLDER, filename, as_attachment=True)


if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=False)
