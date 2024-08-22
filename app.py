from flask import Flask, render_template, request, send_file, after_this_request, flash
from scheduler import InternScheduler
from datetime import datetime
import os

app = Flask(__name__)

# Set a secret key for session management and flash messages
app.secret_key = 'duty_scheduler_chich'.encode('utf8')

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
          interns = [intern.strip() for intern in intern_list]  # Clean up spaces
          units_dict = {intern: units_per_intern for intern in interns}
          scheduler = InternScheduler(start_date, end_date, interns, units_dict, min_interns_per_duty, minimum_spacing)

          # Solve the scheduling problem
          scheduler.solve(randomize=True)

          # Define the output path in the flask_app directory
          output_file = os.path.join(os.getcwd(), "temp", 'schedule.xlsx')

          # Create the calendar and save it as an Excel file
          scheduler.create_calendar(output_file)

          # Hook into the response to delete the file after it has been sent
          @after_this_request
          def remove_file(response):
              try:
                  os.remove(output_file)  # Delete the file after the response is sent
              except Exception as e:
                  print(f"Error removing or closing downloaded file handle: {e}")
              return response

          # Send the Excel file to the user for download
          return send_file(output_file, as_attachment=True)

        except ValueError as e:
              # Catch the ValueError and pass the message to the template
              flash(str(e))
              return render_template('index.html')

        except Exception as e:
            # Catch all other exceptions
            flash(f"An unexpected error occurred: {str(e)}")
            return render_template('index.html')

    return render_template('index.html')



if __name__ == '__main__':
    app.run(debug=True)
