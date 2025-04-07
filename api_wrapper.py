import json
import logging
import traceback
from functools import wraps
from flask import jsonify, Response

def ensure_json_response(f):
    """
    Decorator to ensure that function always returns a valid JSON response.
    If an exception occurs, it will be caught and returned as JSON.
    """
    @wraps(f)
    def decorated_function(*args, **kwargs):
        try:
            # Call the original function
            result = f(*args, **kwargs)
            
            # If it's already a Response, ensure it's JSON
            if isinstance(result, Response):
                # Check if it has the correct content type
                if not result.headers.get('Content-Type', '').startswith('application/json'):
                    # Extract data and convert to proper JSON response
                    data = result.get_data(as_text=True)
                    logging.warning(f"Non-JSON response detected, converting: {data[:100]}...")
                    return jsonify({"success": False, "message": "Non-JSON response detected", "data": data[:500]})
                return result
                
            # If it's a tuple (data, status_code), handle it
            elif isinstance(result, tuple) and len(result) == 2:
                data, status_code = result
                # If data is not a dictionary, wrap it in a dictionary
                if not isinstance(data, dict):
                    data = {"data": data}
                return jsonify(data), status_code
                
            # Otherwise, just jsonify the result
            return jsonify(result)
            
        except Exception as e:
            # Log the error
            logging.error(f"Error in {f.__name__}: {str(e)}")
            logging.error(traceback.format_exc())
            
            # Return a JSON error response
            error_response = {
                "success": False,
                "message": f"An error occurred: {str(e)}",
                "error_type": type(e).__name__
            }
            return jsonify(error_response), 500
            
    return decorated_function

def json_response(success=True, message=None, data=None, status_code=200):
    """
    Helper function to create standardized JSON responses.
    """
    response = {
        "success": success
    }
    
    if message:
        response["message"] = message
        
    if data is not None:
        response["data"] = data
        
    return jsonify(response), status_code
