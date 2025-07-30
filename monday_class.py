import requests
import json
import datetime
import pandas as pd

# ============================================================================
# MONDAY.COM API CONSTANTS AND CONFIGURATION
# ============================================================================
# These are placeholder variables that get populated when creating Board instances
item = ""          # Monday.com item ID (row ID)
columnValue = ""   # Value to set in a column
column = ""        # Column ID to modify
board = ""         # Board ID
apiKey = ""        # Monday.com API key (If you do not include this in your code, you will be prompted to enter it. It can be stored in a json file via the boardGen function)
apiUrl = ""        # Monday.com API endpoint
headers = ""       # HTTP headers for API requests
query = ""         # GraphQL query string

# ============================================================================
# GRAPHQL QUERIES AND MUTATIONS
# ============================================================================

# QUERY: Get board items sorted by ID in descending order. Retrieves all items from a board, ordered by item ID (newest first)
# Returns: Items with their ID, name, and all column values
querySortByID = f'{{ boards(ids: [{board}]) {{ items_page(limit: 500, query_params:  {{ order_by: [ {{column_id:"item_id__1", direction: desc}} ] }} ) {{ cursor items {{ id name column_values {{ text value column {{ title }} }} }} }} }} }}'

# QUERY Columns: Get board columns informationRetrieves all columns from a board with their IDs and titles
# Returns: Column IDs and titles for mapping Excel columns to Monday.com columns
queryColumns = f'{{ boards(ids: {board}) {{ columns {{ id title }} }} }}'

# MUTATE COLUMN: Update a simple column value. Changes the value of a specific column for a specific item
# Parameters: item_id (row), board_id, column_id, new_value
mutateColumn = f'mutation {{ change_simple_column_value(item_id: {item}, board_id: {board}, column_id: {column}, value: "{columnValue}") {{ id }} }}'

# MUTATE ITEM: Create a new item (row) in a board. Adds a new row to the Monday.com board
# Parameters: board_id, item_name, column_values (dict of column_id: value pairs)
createItemMutation = '''
mutation ($boardId: ID!, $itemName: String!, $columnValues: JSON!) {
    create_item (
        board_id: $boardId,
        item_name: $itemName,
        column_values: $columnValues
    ) {
        id
        name
    }
}
'''

# ============================================================================
# CORE API FUNCTIONS
# ============================================================================

def mondayQuery(apiKey, apiUrl, headers, board, query):
    """
    Basic POST request to Monday.com GraphQL API
    
    Args:
        apiKey (str): Monday.com API key
        apiUrl (str): Monday.com API endpoint (usually https://api.monday.com/v2)
        headers (dict): HTTP headers including Authorization
        board (str): Board ID
        query (str): GraphQL query or mutation string
    
    Returns:
        requests.Response: API response object
    """
    data = {'query': query}
    return requests.post(url=apiUrl, json=data, headers=headers)

def importBoardColumns(apiKey, apiUrl, headers, board):
    """
    Creates a dictionary mapping column titles to their IDs and indices

    Args:
        apiKey (str): Monday.com API key
        apiUrl (str): Monday.com API endpoint
        headers (dict): HTTP headers
        board (str): Board ID
    
    Returns:
        dict: Dictionary with column titles as keys and {'id': column_id, 'index': position} as values
        
    """
    columns = mondayQuery(apiKey, apiUrl, headers, board, f'{{ boards(ids: {board}) {{ columns {{ id title }} }} }}')
    columns = columns.json()
    columnDict = {}
    boardColumns = columns['data']['boards'][0]['columns']
    x = -1  # Start at -1 because 'name' column is special
    for column in boardColumns:
        columnDict[column['title']] = {'id': column['id'], 'index': x}
        x+=1
    return columnDict

def boardGen(name, boardID=None, apiKey=None, fileName='boards.json'):
    """
    Board configuration generator and manager
    
    This function manages board configurations by:
    1. Loading existing board configurations from a JSON file
    2. Creating new board configurations if they don't exist
    3. Prompting for API key and board ID if not provided
    4. Saving configurations for future use
    
    Args:
        name (str): Human-readable board name
        boardID (str, optional): Monday.com board ID
        apiKey (str, optional): Monday.com API key
        fileName (str): JSON file to store board configurations
    
    Returns:
        dict: Complete board configuration including properties and columns
    """
    fileDict = {}
    try:
        # Try to load existing board configurations
        with open(fileName, 'r') as json_file:
            data = json.load(json_file)
            fileDict = data
            
            # Check if board with this name already exists
            for board in fileDict['Boards']:
                if fileDict['Boards'][board]['name'] == name:
                    # Use existing configuration
                    boardID = fileDict['Boards'][board]['properties']['board']
                    apiKey = fileDict['Boards'][board]['properties']['apiKey']
                else:
                    # Prompt for missing information
                    if apiKey == None:
                        apiKey = str(input("API Key: "))
                    if boardID == None:
                        boardID = str(input("Board ID: "))   
            
            # Create or update board configuration
            fileDict['Boards'][boardID] = {
                'name': name,
                'id': boardID,
                'properties': {
                    "apiKey": apiKey,
                    "apiUrl": "https://api.monday.com/v2",
                    "headers": {"Authorization": apiKey},
                    "board": boardID
                },
                'columns': importBoardColumns(apiKey, "https://api.monday.com/v2", {"Authorization": apiKey}, boardID)
            }

        # Save updated configuration
        with open(fileName, 'w') as json_file:
            json.dump(fileDict, json_file, indent=4) 
        print(f"Board has been saved to {fileName}")

    except:
        # Handle case where file doesn't exist or board not found
        print("Board not found or file is not created.")
        if apiKey == None:
            apiKey = str(input("API Key: "))
        if boardID == None:
            boardID = str(input("Board ID: "))
        if 'Boards' not in fileDict.keys():
            print("creating boards dict")
            fileDict['Boards'] = {}    
        
        # Create new board configuration
        fileDict['Boards'][boardID] = {
            'name': name,
            'id': boardID,
            'properties': {
                "apiKey": apiKey,
                "apiUrl": "https://api.monday.com/v2",
                "headers": {"Authorization": apiKey},
                "board": boardID
            },
            'columns': importBoardColumns(apiKey, "https://api.monday.com/v2", {"Authorization": apiKey}, boardID)
        }
        
        # Save new configuration
        with open(fileName, 'w') as json_file:
            json.dump(fileDict, json_file, indent=4) 
        print(f"Board has been saved to {fileName}")
    
    return fileDict

# ============================================================================
# MAIN BOARD CLASS
# ============================================================================

class Board:
    """
    Monday.com Board class for managing board operations
    
    Attributes:
        name (str): Human-readable board name
        apiKey (str): Monday.com API key
        apiUrl (str): Monday.com API endpoint
        headers (dict): HTTP headers for API requests
        board (str): Board ID
        columns (dict): Column mapping dictionary
    """
    
    def __init__(self, name, boardID=None, apiKey=None):
        """
        Initialize a Board instance
        
        Args:
            name (str): Board name
            boardID (str, optional): Board ID
            apiKey (str, optional): API key
        """
        key = boardGen(name, boardID, apiKey)
        keyProperties = key['Boards'][boardID]['properties']
        
        self.name = name
        self.fileName = 'key.json'
        self.apiKey = keyProperties['apiKey']
        self.apiUrl = keyProperties['apiUrl']
        self.headers = keyProperties['headers']
        self.board = keyProperties['board']
        self.columns = key['Boards'][boardID]['columns']
    
    def create_item(self, item_name, column_values=None):
        """
        Create a new item (row) in the board
        
        Args:
            item_name (str): Name/title of the new item
            column_values (dict, optional): Dictionary mapping column titles to values
        
        Returns:
            dict: Response from Monday.com API
        """
        if column_values is None:
            column_values = {}
        
        # Convert column titles to column IDs and format values properly
        monday_column_values = {}
        for column_title, value in column_values.items():
            if column_title in self.columns:
                column_id = self.columns[column_title]['id']
                
                # Handle different column types
                if column_id.startswith('date'):  # Date columns
                    try:
                        if pd.notna(value) and str(value).strip():
                            date_str = str(value).strip()
                            
                            # Handle datetime format like "04/30/2025 08:39"
                            if ' ' in date_str:
                                date_str = date_str.split(' ')[0]  # Take only the date part
                            
                            # Convert MM/DD/YYYY to YYYY-MM-DD
                            if '/' in date_str:
                                parts = date_str.split('/')
                                if len(parts) == 3:
                                    month, day, year = parts
                                    formatted_date = f"{year}-{month.zfill(2)}-{day.zfill(2)}"
                                    # Monday.com expects a dict: {"date": "YYYY-MM-DD"}
                                    monday_column_values[column_id] = {"date": formatted_date}
                            else:
                                # Try to parse other date formats
                                monday_column_values[column_id] = {"date": date_str}
                    except Exception as e:
                        print(f"Warning: Could not format date '{value}' for column '{column_title}': {e}")
                        continue
                else:
                    # Regular text/number columns
                    monday_column_values[column_id] = str(value)
        
        # Prepare the mutation
        variables = {
            'boardId': self.board,
            'itemName': item_name,
            'columnValues': json.dumps(monday_column_values)
        }
        
        data = {
            'query': createItemMutation,
            'variables': variables
        }
        
        try:
            response = requests.post(url=self.apiUrl, json=data, headers=self.headers)
            
            # Check if response is successful
            if response.status_code != 200:
                print(f"HTTP Error {response.status_code}: {response.text}")
                return None
            
            # Parse JSON response
            response_data = response.json()
            
            # Check for API errors
            if 'errors' in response_data:
                print(f"API Error: {response_data['errors']}")
                return None
            
            return response_data
            
        except requests.exceptions.RequestException as e:
            print(f"Request failed: {str(e)}")
            return None
        except json.JSONDecodeError as e:
            print(f"JSON decode error: {str(e)}")
            print(f"Response text: {response.text}")
            return None
        except Exception as e:
            print(f"Unexpected error: {str(e)}")
            return None
    
    def update_column_value(self, item_id, column_title, value):
        """
        Update a specific column value for an item
        
        Args:
            item_id (str): Item ID to update
            column_title (str): Column title (must match Monday.com column name)
            value (str): New value to set
        
        Returns:
            dict: Response from Monday.com API
        """
        if column_title not in self.columns:
            raise ValueError(f"Column '{column_title}' not found in board")
        
        column_id = self.columns[column_title]['id']
        mutation = f'mutation {{ change_simple_column_value(item_id: {item_id}, board_id: {self.board}, column_id: {column_id}, value: "{value}") {{ id }} }}'
        
        data = {'query': mutation}
        response = requests.post(url=self.apiUrl, json=data, headers=self.headers)
        return response.json()
    
    def get_items(self, limit=500):
        """
        Get all items from the board
        
        Args:
            limit (int): Maximum number of items to retrieve
        
        Returns:
            dict: Response from Monday.com API
        """
        query = f'{{ boards(ids: [{self.board}]) {{ items_page(limit: {limit}, query_params: {{ order_by: [ {{column_id:"item_id__1", direction: desc}} ] }} ) {{ cursor items {{ id name column_values {{ text value column {{ title }} }} }} }} }} }}'
        
        data = {'query': query}
        response = requests.post(url=self.apiUrl, json=data, headers=self.headers)
        return response.json()
    
    def upload_excel_data(self, excel_file, sheet_name=0, name_column=None):
        """
        Upload data from Excel file to Monday.com board
        
        This method reads an Excel file and creates items in Monday.com for each row.
        It automatically maps Excel column headers to Monday.com columns.
        
        Args:
            excel_file (str): Path to Excel file
            sheet_name (str/int): Sheet name or index (default: first sheet)
            name_column (str, optional): Column to use as item name. If None, uses first column
        
        Returns:
            list: List of created item IDs
        """
        # Read file based on extension
        if excel_file.lower().endswith('.csv'):
            df = pd.read_csv(excel_file)
        else:
            df = pd.read_excel(excel_file, sheet_name=sheet_name)
        
        created_items = []
        
        for index, row in df.iterrows():
            # Determine item name
            if name_column and name_column in row:
                item_name = str(row[name_column])
            else:
                item_name = str(row.iloc[0])  # Use first column as name
            
            # Prepare column values (exclude the name column)
            column_values = {}
            for column_name, value in row.items():
                if column_name != name_column and pd.notna(value):
                    column_values[column_name] = value
            
            # Create item in Monday.com
            try:
                print(f"Creating item {index + 1}/{len(df)}: {item_name}")
                response = self.create_item(item_name, column_values)
                
                if response is None:
                    print(f"Failed to create item: {item_name} - No response from API")
                    continue
                
                if 'data' in response and 'create_item' in response['data']:
                    created_items.append(response['data']['create_item']['id'])
                    print(f"  Created item: {item_name}")
                else:
                    print(f"  Failed to create item: {item_name}")
                    print(f"Response: {response}")
                    
            except Exception as e:
                print(f"  Error creating item {item_name}: {str(e)}")
        
        return created_items
    
    def print_columns(self):
        """
        Print all available columns in the board
        
        Useful for debugging and understanding the board structure
        """
        print(f"\nAvailable columns in board '{self.name}':")
        print("=" * 50)
        for title, info in self.columns.items():
            print(f"{title:<30} | ID: {info['id']:<15} | Index: {info['index']}")
        print("=" * 50) 