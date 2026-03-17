Python code for a streamlit app to modify word documents for improved digital accessibility, using the TAMU AI chat platform for contextual analysis of the document. The app will make the following changes to documents:
  1. Font
  2. Adding a title (if missing)
  3. Identifying titles, heading 1, and heading 2 text, if these structural styles are not used.
  4. Creating structural styles according to user inputs and applying these to the identified text.
  5. Generating alt text for tables and figures.

The user must first enter their TAMU AI chat API key (log in at https://tamus.ai/ and view documentation for directions). Note that the code uses the TAMU API endpoints. Users from other TAMUS entities would need to fork this repo and make the changes to the endpoints in the call_models_api and interact_with_model functions. Users then select which LLM is used to interact with the text and specify font (selection of sans serif fonts) and styles. The code will automatically generate prompts to the LLM; however, the user can optionally add additional instructions for identifying titles and headings in the document, if a specific visual style was used. Multiple word documents or folders can be uploaded.

The modified documents will be returned as a zip file, unless there are three or fewer files, in which case there is an option to download individual files.  

Use Program uses the python-docx library to modify the word documents.
