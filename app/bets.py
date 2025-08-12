from flask import Flask
import papermill as pm

app = Flask(__name__)

@app.route('/refresh-bets')
def run_bets():
    try:
        pm.execute_notebook(
            'extract_info.ipynb',       
            'executed_extract_info.ipynb' 
        )
        return "Notebook executed successfully!"
    except Exception as e:
        return f"Notebook execution failed: {e}", 500

if __name__ == '__main__':
    app.run(port=5000)
