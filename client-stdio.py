# python client-stdio.py
import asyncio
import nest_asyncio
from mcp import ClientSession, StdioServerParameters
from mcp.client.stdio import stdio_client

nest_asyncio.apply()  # Needed to run interactive python


async def main():
    keyword = input("🔍 Enter the word you want to search for in the file Excel: ")
    # Define server parameters
    server_params = StdioServerParameters(
        command="python",  # The command to run your server
        args=["server.py"],  # Arguments to the command
    )

    # Connect to the server
    async with stdio_client(server_params) as (read_stream, write_stream):
        async with ClientSession(read_stream, write_stream) as session:
            # Initialize the connection
            await session.initialize()

            # List available tools
            tools_result = await session.list_tools()
            print("Available tools:")
            for tool in tools_result.tools:
                print(f"  - {tool.name}: {tool.description}")

            # Call our calculator tool
            result = await session.call_tool(
                "fast_search_in_excel",
                arguments={
                    "full_path": r".\excel_file.xlsx",
                    "sheet_name": "Sheet1",  
                    "keyword": keyword,          
                    "usecols": "A:k",            
                    "case_sensitive": False
                }
            )
            print("result:")
            print(result.content[0].text)

if __name__ == "__main__":
    asyncio.run(main())