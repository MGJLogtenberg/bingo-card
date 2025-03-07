from pptx import Presentation
import random
import win32com.client
import os

def read_prompts(prompts_path):
    try:
        with open(prompts_path, 'r', encoding='utf-8') as file:
            return [line.strip() for line in file if line.strip()]
    except Exception as e:
        print(f"Error reading prompts file: {str(e)}")
        return []

def modify_pptx_and_save_pdf(pptx_path, output_pdf_path, phrases):
    try:
        # Open PowerPoint application
        powerpoint = win32com.client.Dispatch("Powerpoint.Application")
        powerpoint.Visible = True  # Keep PowerPoint visible
        
        print("Opening PowerPoint presentation...")
        # Open the presentation
        presentation = powerpoint.Presentations.Open(pptx_path)
        
        # Get the first slide (assuming it's the bingo card)
        slide = presentation.Slides(1)
        
        # Keep track of used phrases to avoid duplicates
        used_phrases = []
        
        # Maximum height in points
        MAX_HEIGHT = 73
        
        print("Replacing text...")
        # Go through each shape in the slide
        for shape in slide.Shapes:
            if shape.HasTextFrame:
                if shape.TextFrame.HasText:
                    text_range = shape.TextFrame.TextRange
                    if text_range.Text.strip().lower() == "test":
                        # Store original position and dimensions
                        original_top = shape.Top
                        original_height = shape.Height
                        
                        # Select a random phrase that hasn't been used yet
                        available_phrases = [p for p in phrases if p not in used_phrases]
                        if not available_phrases:
                            available_phrases = phrases  # Reset if all phrases have been used
                        
                        new_phrase = random.choice(available_phrases)
                        used_phrases.append(new_phrase)
                        
                        print(f"Replacing 'test' with: {new_phrase}")
                        text_range.Text = new_phrase
                        
                        # Set text color to dark orange (RGB: 204, 85, 0)
                        text_range.Font.Color.RGB = 204 + (85 * 256) + (0 * 256 * 256)
                        
                        # Center align the text vertically in the text frame
                        shape.TextFrame.VerticalAnchor = 3  # msoAnchorMiddle
                        
                        # Initial font size
                        initial_font_size = 12
                        text_range.Font.Size = initial_font_size
                        
                        # Check for long words (>11 characters)
                        has_long_word = any(len(word) > 11 for word in new_phrase.split())
                        
                        # If there's a long word, start with a smaller font size
                        if has_long_word:
                            text_range.Font.Size = initial_font_size * 0.8  # 20% smaller
                        
                        # Adjust font size based on width and height constraints
                        while True:
                            # Check text box height and width
                            if shape.Height > MAX_HEIGHT or shape.Width < text_range.BoundWidth:
                                current_size = text_range.Font.Size
                                if current_size <= 6:  # Minimum font size
                                    break
                                text_range.Font.Size = current_size - 0.5
                            else:
                                break
                        
                        # Set the shape height to maximum allowed height
                        shape.Height = MAX_HEIGHT
                        
                        # Move shape up
                        shape.Top = original_top - 35
                        
                        # Ensure text is centered within the shape
                        shape.TextFrame.MarginTop = 0
                        shape.TextFrame.MarginBottom = 0
                        
        print("Saving as PDF...")
        # Save as PDF
        presentation.SaveAs(output_pdf_path, 32)  # 32 is the PDF format code
        
        # Close and clean up
        print("Cleaning up...")
        presentation.Close()
        powerpoint.Quit()
        
        print(f"Successfully created modified PDF at: {output_pdf_path}")
        
    except Exception as e:
        print(f"An error occurred: {str(e)}")
        # Make sure PowerPoint is closed even if there's an error
        try:
            presentation.Close()
            powerpoint.Quit()
        except:
            pass

def create_multiple_bingo_cards(num_cards):
    # File paths
    pptx_path = os.path.abspath(r"..\AICON Bingo.pptx") #Modify paths
    prompts_path = os.path.abspath(r"..\prompts.txt") #Modify paths
    
    try:
        # Read prompts
        print("Reading prompts file...")
        phrases = read_prompts(prompts_path)
        
        if not phrases:
            print("No phrases found in the prompts file.")
            return
        
        # Create directory for multiple cards if it doesn't exist
        output_dir = os.path.abspath(r"..\Bingo_Cards") #Modify paths
        if not os.path.exists(output_dir):
            os.makedirs(output_dir)
        
        # Create multiple bingo cards
        for i in range(num_cards):
            output_pdf_path = os.path.join(output_dir, f"AICON_Bingo_Card_{i+1}.pdf")
            print(f"\nCreating Bingo Card {i+1} of {num_cards}...")
            modify_pptx_and_save_pdf(pptx_path, output_pdf_path, phrases)
            
    except Exception as e:
        print(f"An error occurred in main: {str(e)}")

if __name__ == "__main__":
    # Ask user for number of cards to create
    while True:
        try:
            num_cards = int(input("How many bingo cards would you like to create? "))
            if num_cards > 0:
                break
            else:
                print("Please enter a positive number.")
        except ValueError:
            print("Please enter a valid number.")
    
    create_multiple_bingo_cards(num_cards)
