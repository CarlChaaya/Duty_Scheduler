from datetime import datetime
import os
import secrets
from pathlib import Path

from flask import Flask, render_template, request, send_file, after_this_request, flash

from scheduler import InternScheduler

# Load environment variables from a local .env file if python-dotenv is available.
try:
    from dotenv import load_dotenv  # type: ignore

    load_dotenv()
except ModuleNotFoundError:
    pass

app = Flask(__name__)

# Set a secret key for session management and flash messages
app.config["SECRET_KEY"] = os.environ.get("FLASK_SECRET_KEY", secrets.token_hex(32))

TEMP_DIR = Path(__file__).parent / "temp"
TEMP_DIR.mkdir(exist_ok=True)


@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        try: 
          # Get form data
          start_date = request.form['start_date']
          end_date = request.form['end_date']
          intern_list = request.form['intern_list'].split(',')
          min_interns_per_duty = int(request.form['min_interns_per_duty'])
          units_per_intern = int(request.form['units_per_intern'])
          minimum_spacing = int(request.form['minimum_spacing'])
          
          # Parse dates
          start_date = datetime.strptime(start_date, '%Y-%m-%d')
          end_date = datetime.strptime(end_date, '%Y-%m-%d')

          # Initialize the scheduler
          interns = [intern.strip() for intern in intern_list if intern.strip()]  # Clean up spaces
          if not interns:
              raise ValueError("Please provide at least one intern name.")
          if start_date > end_date:
              raise ValueError("The end date must be on or after the start date.")
          if min_interns_per_duty <= 0:
              raise ValueError("Minimum interns per duty must be greater than zero.")
          if units_per_intern <= 0:
              raise ValueError("Units per intern must be greater than zero.")
          if minimum_spacing < 0:
              raise ValueError("Minimum spacing cannot be negative.")
          units_dict = {intern: units_per_intern for intern in interns}
          scheduler = InternScheduler(start_date, end_date, interns, units_dict, min_interns_per_duty, minimum_spacing)

          # Solve the scheduling problem
          scheduler.solve(randomize=True)

          # Define the output path in the flask_app directory
          output_file = TEMP_DIR / 'schedule.xlsx'

          # Create the calendar and save it as an Excel file
          scheduler.create_calendar(str(output_file))

          # Hook into the response to delete the file after it has been sent
          @after_this_request
          def remove_file(response):
              try:
                  output_file.unlink(missing_ok=True)  # Delete the file after the response is sent
              except Exception as e:
                  app.logger.exception("Error removing or closing downloaded file handle: %s", e)
              return response

          # Send the Excel file to the user for download
          return send_file(output_file, as_attachment=True, download_name="duty-schedule.xlsx")

        except ValueError as e:
              # Catch the ValueError and pass the message to the template
              flash(str(e))
              return render_template('index.html')

        except Exception as e:
            # Catch all other exceptions
            app.logger.exception("Unexpected error while generating schedule")
            flash(f"An unexpected error occurred: {str(e)}")
            return render_template('index.html')

    return render_template('index.html')



if __name__ == '__main__':
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)))
