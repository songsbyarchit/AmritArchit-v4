import openai
from googleapiclient.discovery import build
from google.oauth2 import service_account
import os
from dotenv import load_dotenv
import logging
import requests  # For fetching images
import json      # For parsing Google Custom Search API responses

# Load credentials
SCOPES = ["https://www.googleapis.com/auth/presentations", "https://www.googleapis.com/auth/drive"]
SERVICE_ACCOUNT_FILE = "credentials.json"

credentials = service_account.Credentials.from_service_account_file(
    SERVICE_ACCOUNT_FILE, scopes=SCOPES
)

# Initialize Google APIs
slides_service = build("slides", "v1", credentials=credentials)
drive_service = build("drive", "v3", credentials=credentials)

# Load environment variables
load_dotenv()
OPENAI_API_KEY = os.getenv("OPENAI_API_KEY")
GOOGLE_API_KEY = os.getenv("GOOGLE_API_KEY")  # Add Google API Key
GOOGLE_CSE_ID = os.getenv("GOOGLE_CSE_ID")  # Add 
openai.api_key = OPENAI_API_KEY

def generate_slide_content(topic):
    """Uses OpenAI to generate 5 slide titles and 5 bullet points per slide with hard-hitting, audience-engaging facts."""
    prompt = f"""
    Create a professional PowerPoint presentation on '{topic}' that is engaging, fact-driven, and visually impactful. 
    The slides should be front-facing and ready to present without requiring additional speaker notes.
    
    Instructions:
    - Generate exactly 5 slides.
    - Each slide should have a clear, concise, and compelling title.
    - Provide exactly 5 bullet points per slide that are informative, surprising, or impactful.
    - Avoid generic topic suggestions; instead, include interesting facts, statistics, historical context, or thought-provoking insights.
    - Ensure the language is direct, engaging, and audience-friendly.
    - Format the response with clear slide separations.

    Example:
    Slide 1: [Title]
    - Bullet point 1 (Interesting fact, statistic, or statement)
    - Bullet point 2 (Compelling information)
    - Bullet point 3 (A surprising or little-known fact)
    - Bullet point 4 (Historical or futuristic relevance)
    - Bullet point 5 (Final key takeaway)

    Now, generate the presentation.
    """

    response = openai.ChatCompletion.create(
        model="gpt-4",
        messages=[{"role": "system", "content": "You are an AI that generates structured, engaging, and audience-focused PowerPoint slide content."},
                  {"role": "user", "content": prompt}]
    )

    ai_output = response["choices"][0]["message"]["content"]
    print(f"🤖 OpenAI Response:\n{ai_output}\n")  # Debugging log
    return ai_output

def fetch_image_url(query):
    """Fetches the first image URL from Google Images using Custom Search API."""
    GOOGLE_CSE_ID = os.getenv("GOOGLE_CSE_ID")  # Your Custom Search Engine ID
    GOOGLE_API_KEY = os.getenv("GOOGLE_API_KEY")  # Your Google API Key
    search_url = "https://www.googleapis.com/customsearch/v1"

    params = {
        "q": query,
        "cx": GOOGLE_CSE_ID,
        "key": GOOGLE_API_KEY,
        "searchType": "image",
        "num": 1
    }

    print(f"🔍 Searching image for: {query}")
    response = requests.get(search_url, params=params)
    
    if response.status_code == 200:
        data = response.json()
        if "items" in data and len(data["items"]) > 0:
            image_url = data["items"][0]["link"]
            # Ensure the image is a valid and publicly accessible URL
            if not image_url.startswith("http"):
                print("⚠️ Invalid image URL, skipping.")
                return None
            print(f"✅ Image found: {image_url}")
            return image_url  # Return first image URL
    
    print("❌ No image found.")
    return None

def create_presentation(topic):
    """Creates a new Google Slides presentation."""
    presentation = slides_service.presentations().create(body={"title": f"AI Generated: {topic}"}).execute()
    presentation_id = presentation["presentationId"]
    print(f"✅ Presentation Created: https://docs.google.com/presentation/d/{presentation_id}")
    return presentation_id

def add_slides(presentation_id, slides_content):
    """Adds slides with AI-generated content to Google Slides."""
    requests = []
    
    slides = slides_content.strip().split("\n\n")  # Splitting slides based on spacing
    print(f"📝 Parsed Slides: {slides}\n")  # Debugging log

    for i, slide in enumerate(slides):
        lines = slide.split("\n")
        if len(lines) < 2:
            continue  # Skip malformed slides
        
        slide_title = lines[0].strip()
        bullet_points = "\n".join(lines[1:])

        # Create slide
        create_slide_response = slides_service.presentations().batchUpdate(
            presentationId=presentation_id,
            body={"requests": [{"createSlide": {"slideLayoutReference": {"predefinedLayout": "TITLE_AND_BODY"}}}]}
        ).execute()

        # Get the actual slide ID Google assigned
        slide_id = create_slide_response["replies"][0]["createSlide"]["objectId"]

        # Get placeholders for title and body
        get_page_elements = slides_service.presentations().pages().get(
            presentationId=presentation_id, pageObjectId=slide_id
        ).execute()

        title_id = None
        body_id = None

        for element in get_page_elements.get("pageElements", []):  # Avoids KeyError if empty
            print(f"🔍 Element Found: {element['objectId']} - {element.get('shape', {}).get('placeholder', {}).get('type')}")  # Debugging log

            placeholder_type = element.get("shape", {}).get("placeholder", {}).get("type")

            if placeholder_type == "TITLE":
                title_id = element["objectId"]
            elif placeholder_type == "BODY":
                body_id = element["objectId"]

        print(f"🆔 Found Title ID: {title_id}, Body ID: {body_id}")  # Debugging log

        # Ensure placeholders exist before adding text
        if title_id:
            requests.append({
                "insertText": {
                    "objectId": title_id,
                    "text": slide_title
                }
            })

        if body_id:
            requests.append({
                "insertText": {
                    "objectId": body_id,
                    "text": bullet_points
                }
            })

        # Fetch an image URL for the slide title
        image_url = fetch_image_url(slide_title.split(": ", 1)[-1])  # Remove "Slide X: " prefix
        if image_url and image_url.startswith("http"):
            print(f"🖼️ Adding image to slide '{slide_title}': {image_url}")
            requests.append({
                "createImage": {
                    "url": image_url,
                    "elementProperties": {
                        "pageObjectId": slide_id,
                        "size": {
                            "height": {"magnitude": 3000000, "unit": "EMU"},
                            "width": {"magnitude": 5000000, "unit": "EMU"}
                        },
                        "transform": {
                            "scaleX": 0.7,
                            "scaleY": 0.7,
                            "translateX": 5000000,  # Move image further right
                            "translateY": 500000,   # Adjust vertical alignment
                            "unit": "EMU"
                        }
                    }
                }
            })
        else:
            print(f"⚠️ No image found for slide '{slide_title}'.")

    # Print requests before batchUpdate
    print(f"📤 Sending batch update: {json.dumps(requests, indent=2)}")

    # Only run batchUpdate if there are valid requests
    if requests:
        slides_service.presentations().batchUpdate(presentationId=presentation_id, body={"requests": requests}).execute()
        print("✅ Slides Added!")
    else:
        print("🚨 No valid slide content to add.")

def share_presentation(presentation_id):
    """Shares the presentation with anyone as an editor."""
    permission = {"type": "anyone", "role": "writer"}
    
    drive_service.permissions().create(fileId=presentation_id, body=permission).execute()
    print(f"✅ Shared! Anyone can edit: https://docs.google.com/presentation/d/{presentation_id}")

if __name__ == "__main__":
    topic = input("Enter a topic for the presentation: ")
    slides_content = generate_slide_content(topic)
    presentation_id = create_presentation(topic)
    add_slides(presentation_id, slides_content)
    share_presentation(presentation_id)