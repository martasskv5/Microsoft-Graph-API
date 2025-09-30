def load_env_variables():
    """Load environment variables from the .env file."""
    from dotenv import load_dotenv
    import os

    load_dotenv()

    return {
        "CLIENT_ID": os.getenv("CLIENT_ID"),
        "CLIENT_SECRET": os.getenv("CLIENT_SECRET"),
        "TENANT_ID": os.getenv("TENANT_ID"),
        "REDIRECT_URI": os.getenv("REDIRECT_URI"),
        "SCOPES": os.getenv("SCOPES").split(",") if os.getenv("SCOPES") else []
    }

def handle_error(error):
    """Handle errors by logging or displaying them."""
    print(f"Error: {error}")

def format_output(data):
    """Format output data for display."""
    return "\n".join(f"{key}: {value}" for key, value in data.items())