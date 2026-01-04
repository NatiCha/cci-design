"""
MS Graph client setup with lazy initialization.
"""

from azure.identity import ClientSecretCredential
from msgraph import GraphServiceClient

from core.config import GRAPH_APP_ID, GRAPH_CLIENT_SECRET, GRAPH_TENANT_ID

_graph_client: GraphServiceClient | None = None


def get_graph_client() -> GraphServiceClient:
    """Get or create the MS Graph client (lazy initialization)."""
    global _graph_client
    if _graph_client is None:
        credential = ClientSecretCredential(
            tenant_id=GRAPH_TENANT_ID,
            client_id=GRAPH_APP_ID,
            client_secret=GRAPH_CLIENT_SECRET,
        )
        _graph_client = GraphServiceClient(credentials=credential)
    return _graph_client
