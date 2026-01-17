"""
Create round icon from tools.jpg with drop shadow effect.
Centers on the crossing point of the tools for optimal icon appearance.
"""

from PIL import Image, ImageDraw, ImageFilter
import os
import numpy as np

script_dir = os.path.dirname(os.path.abspath(__file__))
input_image = os.path.join(script_dir, 'tools.jpg')
output_ico = os.path.join(script_dir, 'myicon.ico')
output_png = os.path.join(script_dir, 'myicon.png')

# Theme colors (matching outlook_miner_qt.py)
DARK_SLATE = (44, 62, 80)      # #2C3E50 - header_bg


def find_crossing_center(is_tools):
    """Find the center point where the tools cross."""
    # The crossing point is typically where tool pixels are densest
    # We'll find the centroid of the tool pixels
    tool_coords = np.where(is_tools)
    if len(tool_coords[0]) > 0:
        center_row = int(np.mean(tool_coords[0]))
        center_col = int(np.mean(tool_coords[1]))
        return center_row, center_col
    return None, None


def create_icon_from_tools(size):
    """Create round icon from tools.jpg with white tools on slate background."""
    # Load the tools image
    tools_img = Image.open(input_image)

    # Convert to RGBA
    if tools_img.mode != 'RGBA':
        tools_img = tools_img.convert('RGBA')

    # Get the image data as numpy array for processing
    img_array = np.array(tools_img)

    r, g, b = img_array[:, :, 0], img_array[:, :, 1], img_array[:, :, 2]

    # Pixels are considered background if they're light (R, G, B all > 200)
    is_background = (r > 200) & (g > 200) & (b > 200)

    # Pixels are considered tools if they're dark (not background)
    is_tools = ~is_background

    # Find the crossing center of the tools
    center_row, center_col = find_crossing_center(is_tools)

    # Find the extent of the tools from the center
    tool_coords = np.where(is_tools)
    min_row, max_row = tool_coords[0].min(), tool_coords[0].max()
    min_col, max_col = tool_coords[1].min(), tool_coords[1].max()

    # Calculate the maximum distance from center to any edge of the tools
    dist_top = center_row - min_row
    dist_bottom = max_row - center_row
    dist_left = center_col - min_col
    dist_right = max_col - center_col

    # Use the maximum distance to create a square crop centered on the crossing
    max_dist = max(dist_top, dist_bottom, dist_left, dist_right)

    # Add padding around the tools (15% of max distance)
    padding = int(max_dist * 0.15)
    crop_radius = max_dist + padding

    # Calculate crop boundaries centered on the crossing point
    crop_top = max(0, center_row - crop_radius)
    crop_bottom = min(img_array.shape[0], center_row + crop_radius)
    crop_left = max(0, center_col - crop_radius)
    crop_right = min(img_array.shape[1], center_col + crop_radius)

    # Make sure crop is square
    crop_height = crop_bottom - crop_top
    crop_width = crop_right - crop_left
    if crop_height != crop_width:
        min_dim = min(crop_height, crop_width)
        crop_bottom = crop_top + min_dim
        crop_right = crop_left + min_dim

    # Crop the original image centered on crossing
    cropped = tools_img.crop((crop_left, crop_top, crop_right, crop_bottom))
    cropped_array = np.array(cropped)

    # Recompute is_background and is_tools for cropped image
    r_c = cropped_array[:, :, 0]
    g_c = cropped_array[:, :, 1]
    b_c = cropped_array[:, :, 2]
    is_background_c = (r_c > 200) & (g_c > 200) & (b_c > 200)
    is_tools_c = ~is_background_c

    # Replace background with dark slate
    cropped_array[is_background_c, 0] = DARK_SLATE[0]
    cropped_array[is_background_c, 1] = DARK_SLATE[1]
    cropped_array[is_background_c, 2] = DARK_SLATE[2]
    cropped_array[is_background_c, 3] = 255

    # Replace tools with white
    cropped_array[is_tools_c, 0] = 255
    cropped_array[is_tools_c, 1] = 255
    cropped_array[is_tools_c, 2] = 255
    cropped_array[is_tools_c, 3] = 255

    # Convert back to PIL Image
    tools_processed = Image.fromarray(cropped_array)

    # Determine icon parameters based on size
    if size <= 32:
        # Very small icons - fill the space, no shadow
        circle_padding = 1
        shadow_offset = 0
        shadow_blur = 0
    elif size <= 48:
        # Small icons - minimal padding, no shadow
        circle_padding = 2
        shadow_offset = 0
        shadow_blur = 0
    elif size <= 64:
        # Medium icons - slight padding, subtle shadow
        circle_padding = 3
        shadow_offset = 2
        shadow_blur = 2
    else:
        # Large icons - proper padding and shadow for floating effect
        circle_padding = max(4, size // 32)
        shadow_offset = max(3, size // 64)
        shadow_blur = max(4, size // 50)

    # Calculate circle diameter
    if shadow_offset > 0:
        circle_diameter = size - (circle_padding * 2) - shadow_offset
    else:
        circle_diameter = size - (circle_padding * 2)

    # Resize tools image to fit inside the circle with some internal padding
    internal_padding = max(2, circle_diameter // 10)  # 10% internal padding
    tools_size = circle_diameter - (internal_padding * 2)
    tools_resized = tools_processed.resize((tools_size, tools_size), Image.Resampling.LANCZOS)

    # Create the circular background
    circle_img = Image.new('RGBA', (circle_diameter, circle_diameter), (0, 0, 0, 0))
    circle_draw = ImageDraw.Draw(circle_img)
    circle_draw.ellipse((0, 0, circle_diameter - 1, circle_diameter - 1), fill=DARK_SLATE + (255,))

    # Paste the tools on the circle - shift up slightly for visual centering
    # The tools have more visual weight at top (heads) so move up for better visual balance
    visual_offset_y = -max(2, internal_padding // 2)
    tools_offset_x = internal_padding
    tools_offset_y = internal_padding + visual_offset_y
    circle_img.paste(tools_resized, (tools_offset_x, tools_offset_y), tools_resized.split()[3] if tools_resized.mode == 'RGBA' else None)

    # Create circular mask for the final circle
    circle_mask = Image.new('L', (circle_diameter, circle_diameter), 0)
    mask_draw = ImageDraw.Draw(circle_mask)
    mask_draw.ellipse((0, 0, circle_diameter - 1, circle_diameter - 1), fill=255)

    # Apply circular mask
    circle_img.putalpha(circle_mask)

    # Create final image
    result_img = Image.new('RGBA', (size, size), (0, 0, 0, 0))

    if shadow_offset > 0:
        # Create shadow
        shadow = Image.new('RGBA', (size, size), (0, 0, 0, 0))
        shadow_x = circle_padding + shadow_offset
        shadow_y = circle_padding + shadow_offset

        # Draw shadow circle
        shadow_draw = ImageDraw.Draw(shadow)
        shadow_draw.ellipse(
            (shadow_x, shadow_y, shadow_x + circle_diameter - 1, shadow_y + circle_diameter - 1),
            fill=(0, 0, 0, 80)
        )

        # Blur the shadow
        shadow = shadow.filter(ImageFilter.GaussianBlur(shadow_blur))

        # Composite shadow
        result_img = Image.alpha_composite(result_img, shadow)

        # Paste circle (positioned slightly up-left of shadow)
        result_img.paste(circle_img, (circle_padding, circle_padding), circle_img)
    else:
        # No shadow - just center the circle
        result_img.paste(circle_img, (circle_padding, circle_padding), circle_img)

    return result_img


if __name__ == '__main__':
    if not os.path.exists(input_image):
        print(f"Error: {input_image} not found")
        exit(1)

    # Create icons at multiple sizes for ICO file
    sizes = [16, 32, 48, 64, 128, 256]
    icons = [create_icon_from_tools(size) for size in sizes]

    # Save as ICO with multiple sizes (largest first for best quality)
    icons[-1].save(output_ico, format='ICO', sizes=[(s, s) for s in sizes], append_images=icons[:-1])
    print(f"Created {output_ico} with sizes: {sizes}")

    # Save 64x64 PNG for window icon
    png_icon = create_icon_from_tools(64)
    png_icon.save(output_png, format='PNG')
    print(f"Created {output_png} at 64x64 pixels")

    print("\nIcon created from tools.jpg:")
    print("  - Centered on tool crossing point")
    print("  - White tools on dark slate background (#2C3E50)")
    print("  - Circular shape with appropriate padding")
    print("  - Drop shadow for larger sizes")
