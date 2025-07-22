from mcp import ClientSession
from mcp.client.streamable_http import streamablehttp_client

from langgraph.prebuilt import create_react_agent
from langchain_mcp_adapters.tools import load_mcp_tools

import asyncio
from mcp.client.streamable_http import streamablehttp_client
from mcp.client.session import ClientSession
from dotenv import load_dotenv

load_dotenv()
async def main():
    async with streamablehttp_client("https://mcp-excel-finder-deployment.onrender.com/mcp") as (read, write, _):
        async with ClientSession(read, write) as session:
            await session.initialize()
            tools = await load_mcp_tools(session)
            for tool in tools:
                print(f"{tool.name}: {tool.description}")
            agent = create_react_agent("openai:gpt-4.1-mini", tools)
            question = input("Enter your question: ")
            response = await agent.ainvoke({"messages": [("user", question)]})
            final_content = response["messages"][-1].content
            print(final_content)
            
if __name__ == "__main__":
    asyncio.run(main())
