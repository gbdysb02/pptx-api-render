from flask import Flask, request, send_file
from generate_pptx import generate_pptx
import json
import os

app = Flask(__name__)

@app.route("/generate", methods=["POST"])
def generate():
    try:
        companies = request.json.get("companies", [])
        if not companies:
            return {"error": "No companies provided"}, 400

        output_file = "Company_Summary_Deck_Full.pptx"
        print("Received POST with companies:", companies)
        generate_pptx(companies, output_file)
        return send_file(output_file, as_attachment=True)
    except Exception as e:
        return {"error": str(e)}, 500

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=8080)
