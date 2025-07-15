# A Python client to interact with the ms-365-mcp-server and an OpenAI-compatible LLM.
#
# Installation:
#   pip install fastmcp httpx
#
# Instructions:
# 1. Start the server with verbose logging and auth tools enabled:
#    node dist/index.js --http 3000 --enable-auth-tools -v
# 2. Run this client script:
#    python client.py

import asyncio
import json
import os
import sys
import httpx
from fastmcp import Client

# --- No Proxy Setup ---
for var in ("HTTP_PROXY", "http_proxy", "HTTPS_PROXY", "https_proxy"):
    os.environ.pop(var, None)

# --- Configuration ---
SERVER_URL = "http://localhost:3000/mcp"

# --- OpenAI/Bedrock API Configuration ---
BEDROCK_API_BASE_URL = "https://bedrock-api-d.inbison.com/api/"
BEDROCK_API_KEY = "A2@r!2o#@112"
BEDROCK_MODEL_ID = "us.anthropic.claude-3-7-sonnet-20250219-v1:0"

# --- System Prompt ---
SYSTEM_PROMPT = """
You are a helpful assistant that can connect to Microsoft 365.
You must first call the 'connect_and_authenticate' tool.
Once connected, you can use other Microsoft 365 tools to help the user.
Always inform the user about the steps you are taking.
"""

# --- LLM Client Initialization ---
# WARNING: Disabling SSL verification is a security risk. Use only in trusted environments.
limits = httpx.Limits(max_connections=10, max_keepalive_connections=5)
transport = httpx.AsyncHTTPTransport(retries=3, verify=False)
http_client = httpx.AsyncClient(transport=transport, timeout=60.0, limits=limits)

def get_m365_tools():
    """Returns a comprehensive, static list of tool specifications for the LLM."""
    return [
        {"type": "function", "function": {"name": "list-mail-messages", "description": "Lists email messages in the user's mailbox."}},
        {"type": "function", "function": {"name": "list-mail-folders", "description": "Lists all mail folders in the user's mailbox."}},
        {"type": "function", "function": {"name": "list-mail-folder-messages", "description": "Lists messages in a specific mail folder.", "parameters": {"type": "object", "properties": {"mailFolder_id": {"type": "string", "description": "The ID of the mail folder."}}, "required": ["mailFolder_id"]}}},
        {"type": "function", "function": {"name": "get-mail-message", "description": "Gets a specific email message by its ID.", "parameters": {"type": "object", "properties": {"message_id": {"type": "string", "description": "The ID of the message."}}, "required": ["message_id"]}}},
        {"type": "function", "function": {"name": "send-mail", "description": "Sends an email message.", "parameters": {"type": "object", "properties": {"Message": {"type": "object", "properties": {"toRecipients": {"type": "array", "items": {"type": "object", "properties": {"emailAddress": {"type": "object", "properties": {"address": {"type": "string"}, "name": {"type": "string"}}}}}}, "subject": {"type": "string"}, "body": {"type": "object", "properties": {"contentType": {"type": "string", "enum": ["Text", "HTML"]}, "content": {"type": "string"}}}}}, "SaveToSentItems": {"type": "boolean"}}, "required": ["Message"]}}},
        {"type": "function", "function": {"name": "create-calendar-event", "description": "Creates a new event in the user's calendar.", "parameters": {"type": "object", "properties": {"subject": {"type": "string"}, "start": {"type": "object", "properties": {"dateTime": {"type": "string", "description": "ISO 8601 format, e.g., '2025-07-15T10:00:00'"}, "timeZone": {"type": "string", "description": "e.g., 'Pacific Standard Time'"}}}, "end": {"type": "object", "properties": {"dateTime": {"type": "string", "description": "End time in ISO 8601 format"}, "timeZone": {"type": "string"}}}}, "required": ["subject", "start", "end"]}}},
        {"type": "function", "function": {"name": "get-current-user", "description": "Gets information about the currently authenticated user."}},
    ]

class M365Agent:
    """Manages the connection and interaction with the M365 MCP server."""
    def __init__(self, server_url):
        self.server_url = server_url
        self.mcp_client = None
        self._client_context = None

    async def connect_and_authenticate(self):
        """Connects to the server and authenticates the user."""
        if self.mcp_client:
            return "Already connected and authenticated."
        try:
            print(f"[AGENT] Establishing connection to MCP server at {self.server_url}...")
            self._client_context = Client(self.server_url)
            self.mcp_client = await self._client_context.__aenter__()
            print("[AGENT] MCP Client connected successfully.")

            print('\n[AGENT] --- Initiating Login ---')
            # The tool name is 'login'. fastmcp creates a method with this name.
            login_result = await self.mcp_client.login(force=True)
            
            if login_result and login_result.get('content') and login_result['content'][0].get('text'):
                login_message = login_result['content'][0]['text']
                print("\n" + "="*50)
                print("[SERVER AUTHENTICATION MESSAGE]")
                print(login_message)
                print("="*50 + "\n")
            else:
                raise Exception("Did not receive the expected login message from the server.")

            input('[AGENT] --- Press Enter after you have completed the browser login. ---')

            # The tool name is 'verify-login', which becomes 'verify_login' in Python
            verify_result = await self.mcp_client.verify_login()
            if not (verify_result and verify_result.get('content') and verify_result['content'][0].get('text')):
                raise Exception("Did not receive a valid verification response from the server.")

            verification = json.loads(verify_result['content'][0]['text'])
            if verification.get('success'):
                user_name = verification.get('userData', {}).get('displayName', 'Unknown User')
                print(f"[AGENT] Login verification successful for: {user_name}")
                return f"Successfully connected and authenticated to Microsoft 365 as {user_name}."
            else:
                raise Exception(f"Login verification failed: {verification.get('message', 'No message.')}")
        except Exception as e:
            print(f"[AGENT] Error during connection/authentication: {e}")
            await self.disconnect()
            return f"Failed to connect and authenticate. Error: {e}"

    async def execute_tool(self, tool_name, args):
        """Executes a given tool with the provided arguments."""
        if not self.mcp_client:
            return "Error: Not connected. Please call 'connect_and_authenticate' first."
        
        print(f"[AGENT] Executing tool: {tool_name} with args: {args}")
        try:
            tool_function_name = tool_name.replace('-', '_')
            tool_function = getattr(self.mcp_client, tool_function_name)
            function_response = await tool_function(**args)
            tool_output_text = function_response.get('content', [{}])[0].get('text', '')
            print(f"[AGENT] Tool '{tool_name}' response received.")
            return tool_output_text
        except AttributeError:
            print(f"[AGENT] Error: Tool '{tool_name}' (as '{tool_function_name}') not found on mcp_client.")
            return f"Error: The tool '{tool_name}' is not available."
        except Exception as e:
            print(f"[AGENT] Error calling tool '{tool_name}': {e}")
            return f"Error executing tool: {e}"

    async def disconnect(self):
        """Disconnects from the MCP server."""
        if self._client_context and self.mcp_client:
            try:
                await self._client_context.__aexit__(None, None, None)
                print("[AGENT] Disconnected from MCP server.")
            except Exception as e:
                print(f"[AGENT] Error during disconnect: {e}")
        self.mcp_client = None
        self._client_context = None

async def test_llm_connection():
    """Tests connectivity to the Bedrock API."""
    print("\n[LLM] --- Testing Bedrock API Connection ---")
    chat_endpoint = f"{BEDROCK_API_BASE_URL.strip('/')}/v1/chat/completions"
    headers = {"Authorization": f"Bearer {BEDROCK_API_KEY}", "Content-Type": "application/json"}
    payload = {"model": BEDROCK_MODEL_ID, "messages": [{"role": "user", "content": "Health check"}], "max_tokens": 2}
    try:
        response = await http_client.post(chat_endpoint, headers=headers, json=payload)
        response.raise_for_status()
        print("[LLM] Bedrock API connection test successful (received a 2xx response).")
        return True
    except httpx.RequestError as e:
        print(f"\n[LLM ERROR] A connection error occurred: {e}")
        return False
    except httpx.HTTPStatusError as e:
        print(f"\n[LLM ERROR] An HTTP error occurred: {e.response.status_code}\n  Response Body: {e.response.text}")
        return False

async def chat_loop(agent: M365Agent):
    """Handles the conversational interaction with the LLM and the M365 agent."""
    print("\n[CLIENT] --- Starting Chat Session ---\nType 'exit' to end the session.")
    conversation_history = [{"role": "system", "content": SYSTEM_PROMPT}]

    while True:
        try:
            user_input = input("\nYou: ")
            if user_input.lower() == 'exit':
                break

            conversation_history.append({"role": "user", "content": user_input})

            tools_for_llm = []
            if not agent.mcp_client:
                tools_for_llm.append({
                    "type": "function",
                    "function": {
                        "name": "connect_and_authenticate",
                        "description": "Connects to Microsoft 365. Must be called before other tools.",
                        "parameters": {"type": "object", "properties": {}}
                    }
                })
            else:
                tools_for_llm = get_m365_tools()
            
            chat_endpoint = f"{BEDROCK_API_BASE_URL.strip('/')}/v1/chat/completions"
            headers = {"Authorization": f"Bearer {BEDROCK_API_KEY}", "Content-Type": "application/json"}
            payload = {"model": BEDROCK_MODEL_ID, "messages": conversation_history, "tools": tools_for_llm, "tool_choice": "auto"}

            print(f"\n[LLM] ---> Sending request to Bedrock...")
            response = await http_client.post(chat_endpoint, headers=headers, json=payload)
            response.raise_for_status()
            response_data = response.json()
            response_message = response_data['choices'][0]['message']
            print("[LLM] <--- Received response from Bedrock.")

            if response_message.get('tool_calls'):
                print(f"[ASSISTANT] Decided to use a tool: {json.dumps(response_message['tool_calls'])}")
                conversation_history.append(response_message)
                
                for tool_call in response_message['tool_calls']:
                    function_name = tool_call['function']['name']
                    function_args = json.loads(tool_call['function']['arguments'])
                    
                    if function_name == 'connect_and_authenticate':
                        tool_output = await agent.connect_and_authenticate()
                    else:
                        tool_output = await agent.execute_tool(function_name, function_args)
                    
                    conversation_history.append({"tool_call_id": tool_call['id'], "role": "tool", "name": function_name, "content": tool_output})

                print("\n[LLM] ---> Sending tool results back for a final answer...")
                payload.pop("tools", None)
                payload.pop("tool_choice", None)
                second_response = await http_client.post(chat_endpoint, headers=headers, json=payload)
                second_response.raise_for_status()
                final_message = second_response.json()['choices'][0]['message']
                print("[LLM] <--- Received final response.")
                print(f"\nAssistant: {final_message['content']}")
                conversation_history.append(final_message)
            else:
                assistant_response = response_message['content']
                print(f"\nAssistant: {assistant_response}")
                conversation_history.append({"role": "assistant", "content": assistant_response})

        except httpx.HTTPStatusError as e:
            print(f"\n[LLM ERROR] HTTP {e.response.status_code}: {e.response.text}")
            if conversation_history and conversation_history[-1]["role"] == "user":
                conversation_history.pop()
        except Exception as e:
            print(f"\n[CLIENT] An unexpected error occurred: {e}")
            if conversation_history and conversation_history[-1]["role"] == "user":
                conversation_history.pop()

async def main():
    """Initializes the agent and starts the chat loop."""
    if not await test_llm_connection():
        sys.exit(1)

    agent = M365Agent(SERVER_URL)
    try:
        await chat_loop(agent)
    except Exception as e:
        print(f"[CLIENT] A critical error occurred: {e}")
    finally:
        await agent.disconnect()

if __name__ == "__main__":
    asyncio.run(main())
