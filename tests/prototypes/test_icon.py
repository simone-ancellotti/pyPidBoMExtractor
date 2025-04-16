from PIL import Image, ImageDraw

def create_valve_icon(output_path="bom_valve_icon.png"):
    # Define icon dimensions
    img_size = (128, 128)
    # Create a new image with RGBA mode and a transparent background
    img = Image.new("RGBA", img_size, (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)

    # Draw the circular background
    circle_color = (30, 144, 255, 255)  # Dodger Blue
    margin = 5  # margin from the edge for the circle
    ellipse_bbox = [margin, margin, img_size[0]-margin, img_size[1]-margin]
    draw.ellipse(ellipse_bbox, fill=circle_color)

    # Draw the valve symbol as a white diamond in the center.
    center_x, center_y = img_size[0] // 2, img_size[1] // 2
    diamond_half = 40  # half the width/height of the diamond
    diamond_coords = [
        (center_x, center_y - diamond_half),  # top
        (center_x + diamond_half, center_y),  # right
        (center_x, center_y + diamond_half),  # bottom
        (center_x - diamond_half, center_y)   # left
    ]
    valve_color = (255, 255, 255, 255)  # White
    draw.polygon(diamond_coords, fill=valve_color)

    # Optionally, add an inner circle to simulate a valve stem or accent.
    inner_radius = 10
    inner_bbox = [
        center_x - inner_radius, center_y - inner_radius,
        center_x + inner_radius, center_y + inner_radius
    ]
    # Use the same blue as the background to create contrast.
    draw.ellipse(inner_bbox, fill=circle_color)

    # Draw a border around the outer circle.
    border_color = (0, 0, 0, 255)  # Black
    draw.ellipse(ellipse_bbox, outline=border_color, width=3)

    # Save the icon image as a PNG file.
    img.save(output_path, format="PNG")
    print(f"Valve icon created and saved as {output_path}")

if __name__ == "__main__":
    create_valve_icon()
