# app.py
from flask import Flask, render_template, request

app = Flask(__name__)

@app.route('/', methods=['GET', 'POST'])
def company_form():
    if request.method == 'POST':
        # Extract form data
        company_name = request.form.get('company_name')
        market_value = request.form.get('market_value')
        location = request.form.get('location')
        founder_name = request.form.get('founder_name')
        business_type = request.form.get('business_type')

        # Here you could add code to process the data, save to a database, etc.

        # For now, let's just display the received data in the console
        print(f"Company Name: {company_name}")
        print(f"Market Value: {market_value}")
        print(f"Location: {location}")
        print(f"Founder Name: {founder_name}")
        print(f"Business Type: {business_type}")

        return f"Form submitted! Company Name: {company_name}, Market Value: {market_value}"

    return render_template('index.html')

if __name__ == '__main__':
    app.run(debug=True)
