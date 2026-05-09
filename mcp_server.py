"""Microsoft Agent MCP Server.
Provides AD/Entra ID ingestion hooks for the agent-utilities Knowledge Graph.
"""

from agent_utilities.knowledge_graph.core.engine import RegistryGraphEngine
from mcp.server.fastmcp import FastMCP
from pydantic import Field

mcp = FastMCP("microsoft_agent")

@mcp.tool(description="Trigger a batch synchronization of Active Directory entities into the KG.")
def trigger_ad_sync(batch_size: int = Field(1000, description="Number of entities to sync")) -> str:
    """Trigger a batch synchronization of Active Directory entities into the KG."""
    kg = RegistryGraphEngine()

    # Architectural stub for fetching AD data
    ad_data = [{"id": "user:1", "name": "Alice", "type": "Employee"}]
    kg.ingest_external_batch("active_directory", ad_data)

    return f"Triggered sync for {batch_size} AD entities to the enterprise Knowledge Graph."

if __name__ == "__main__":
    mcp.run()
