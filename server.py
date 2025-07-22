# uv run server.py
# npx @modelcontextprotocol/inspector

from mcp.server.fastmcp import FastMCP
from typing import Union
import pandas as pd
import os

# Create an MCP server
mcp = FastMCP("MCP_excel", host="0.0.0.0", port=8000)
excel_file_path = os.path.join(os.path.dirname(__file__), "excel_file.xlsx")

@mcp.tool()
def fast_search_in_excel(
    keyword: Union[str, int],
    case_sensitive: bool = False
) -> str:
    """
    Search for a keyword in 'Sheet1' of an Excel sheet in columns A:K and return matching rows in JSON format.

    Args:
        keyword: The keyword (string or integer) to search for in the cells.
        case_sensitive: If True, performs case-sensitive search (default: False).

    Returns:
        JSON string of matching rows or a message if no matches found.
        Error message if an exception occurs.
    """

    def _prepare_search_string(value: Union[str, int]) -> str:
        """Convert value to string and handle case sensitivity."""
        search_str = str(value)
        return search_str if case_sensitive else search_str.lower()
    
    def _row_matches(row: pd.Series, search_str: str) -> bool:
        """Check if any cell in the row contains the search string."""
        for cell in row:
            cell_str = str(cell)
            if not case_sensitive:
                cell_str = cell_str.lower()
            if search_str in cell_str:
                return True
        return False

    try:
        # Load the Excel sheet
        df = pd.read_excel(
            excel_file_path,
            sheet_name="Sheet1",
            usecols="A:K",  
            dtype=str,
            engine='openpyxl'
        )
        
        # Prepare search string
        search_str = _prepare_search_string(keyword)
        
        # Find matching rows
        mask = df.apply(_row_matches, axis=1, search_str=search_str)
        matched_rows = df[mask]
        
        # Return results
        if matched_rows.empty:
            return "No matching rows found."
            
        return matched_rows.to_json(
            orient="records",
            force_ascii=False,
            indent=2
        )
        
    except FileNotFoundError:
        return f"Error: File not found at path '{excel_file_path}'"
    except ValueError as ve:
        return f"Error: {str(ve)}"
    except Exception as e:
        return f"Unexpected error: {str(e)}"
    
# Run the server
if __name__ == "__main__":
     mcp.run(transport="streamable-http")
