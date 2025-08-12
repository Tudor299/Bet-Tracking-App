from flask import Flask
import papermill as pm

app = Flask(__name__)

@app.route('/refresh-TM-Data')
def run_teams():
    try:
        pm.execute_notebook(
            'TM_Data.ipynb',       
            'executed_TM_Data.ipynb' 
        )
        return "Notebook executed successfully!"
    except Exception as e:
        return f"Notebook execution failed: {e}", 500

if __name__ == '__main__':
    app.run(port=5001)
