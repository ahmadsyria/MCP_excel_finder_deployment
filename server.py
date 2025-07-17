# uv run server.py
from mcp.server.fastmcp import FastMCP
from typing import Optional, Union
import pandas as pd

mcp = FastMCP("DemoServer", host="127.0.0.1", port=8050)

@mcp.tool()
def fast_search_in_excel(
    full_path,
    sheet_name: str,
    keyword: Union[str, int],
    usecols: Optional[str] = None,
    case_sensitive: bool = False
) -> str:
    """
    Search for a keyword in an Excel sheet and return matching rows in JSON format.

    Args:
        sheet_name: Name of the sheet to search within.
        keyword: The keyword (string or integer) to search for in the cells.
        usecols: Optional column range to limit the search (e.g., "A:C").
        case_sensitive: If True, performs case-sensitive search (default: False).

    Returns:
        JSON string of matching rows or a message if no matches found.
        Error message if an exception occurs.

    Example:
        >>> fast_search_in_excel("Sheet1", "example", "A:C")
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
            full_path,
            sheet_name=sheet_name,
            usecols=usecols,
            dtype=str,
            engine='openpyxl'  # Explicit engine for better compatibility
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
        return f"Error: File not found at path '{full_path}'"
    except ValueError as ve:
        return f"Error: {str(ve)}"
    except Exception as e:
        return f"Unexpected error: {str(e)}"

# Run the server
if __name__ == "__main__":
     transport = "stdio"
     print("Running server with stdio transport")
     mcp.run(transport="stdio")
