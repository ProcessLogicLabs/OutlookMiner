"""
Create DocuShuttle icon - Dynamic envelope being pulled into a vortex.
"""

from PIL import Image, ImageDraw, ImageFilter
import os
import math

script_dir = os.path.dirname(os.path.abspath(__file__))
output_ico = os.path.join(script_dir, 'myicon.ico')
output_png = os.path.join(script_dir, 'myicon.png')

# Theme colors
TEAL = (0, 161, 156)              # #00A19C - primary
TEAL_DARK = (0, 120, 116)         # Darker teal
TEAL_LIGHT = (80, 200, 195)       # Lighter teal
WHITE = (255, 255, 255)


def create_icon(size):
    """Create dynamic DocuMiner icon with envelope in vortex."""
    img = Image.new('RGBA', (size, size), (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)

    padding = max(1, size // 32)
    circle_size = size - padding * 2
    center = size // 2
    radius = circle_size // 2

    # Draw teal background
    draw.ellipse(
        [padding, padding, padding + circle_size - 1, padding + circle_size - 1],
        fill=TEAL
    )

    # Draw dynamic vortex swirl - multiple curved streaks
    for i in range(6):
        base_angle = i * 60

        # Each swirl arm
        for j in range(30):
            t = j / 29
            # Spiral equation
            angle = base_angle + t * 280
            dist = 6 + t * (radius - 10)

            x = center + dist * math.cos(math.radians(angle))
            y = center + dist * math.sin(math.radians(angle))

            # Streak size varies
            streak_size = max(1, int((3 - t * 2) * (size / 64)))

            # Color gets lighter toward outside
            brightness = int(80 + t * 60)
            alpha = int(200 - t * 100)
            color = (brightness, 220 - int(t * 40), 215 - int(t * 40), alpha)

            if streak_size > 0:
                draw.ellipse(
                    [x - streak_size, y - streak_size, x + streak_size, y + streak_size],
                    fill=color
                )

    # Draw glowing vortex center
    glow_radius = max(6, size // 8)
    for r in range(glow_radius, 0, -1):
        t = 1 - (r / glow_radius)
        # White to light teal gradient
        color = (
            int(TEAL_LIGHT[0] + (255 - TEAL_LIGHT[0]) * t),
            int(TEAL_LIGHT[1] + (255 - TEAL_LIGHT[1]) * t),
            int(TEAL_LIGHT[2] + (255 - TEAL_LIGHT[2]) * t),
            int(150 + 105 * t)
        )
        draw.ellipse(
            [center - r, center - r, center + r, center + r],
            fill=color
        )

    # Draw envelope - tilted as if being pulled in
    env_width = int(circle_size * 0.48)
    env_height = int(env_width * 0.65)

    # Position slightly off-center (being pulled toward vortex)
    env_center_x = center + size // 20
    env_center_y = center - size // 16

    # Calculate envelope corners with rotation
    rotation = -15  # Tilted
    rot_rad = math.radians(rotation)

    def rotate_point(px, py, cx, cy, angle):
        cos_a = math.cos(angle)
        sin_a = math.sin(angle)
        dx = px - cx
        dy = py - cy
        return (cx + dx * cos_a - dy * sin_a, cy + dx * sin_a + dy * cos_a)

    # Envelope base corners
    half_w = env_width // 2
    half_h = env_height // 2

    corners = [
        (env_center_x - half_w, env_center_y - half_h),  # Top left
        (env_center_x + half_w, env_center_y - half_h),  # Top right
        (env_center_x + half_w, env_center_y + half_h),  # Bottom right
        (env_center_x - half_w, env_center_y + half_h),  # Bottom left
    ]

    # Rotate all corners
    rotated = [rotate_point(p[0], p[1], env_center_x, env_center_y, rot_rad) for p in corners]

    # Draw envelope shadow
    shadow_offset = max(2, size // 32)
    shadow_corners = [(p[0] + shadow_offset, p[1] + shadow_offset) for p in rotated]
    draw.polygon(shadow_corners, fill=(0, 80, 78, 80))

    # Draw envelope body
    draw.polygon(rotated, fill=WHITE)

    # Flap triangle
    flap_tip = rotate_point(env_center_x, env_center_y - half_h + env_height * 0.38,
                            env_center_x, env_center_y, rot_rad)
    flap_points = [rotated[0], flap_tip, rotated[1]]
    draw.polygon(flap_points, fill=(235, 235, 235))

    # Flap lines
    line_width = max(1, size // 64)
    draw.line([rotated[0], flap_tip], fill=(200, 200, 200), width=line_width)
    draw.line([rotated[1], flap_tip], fill=(200, 200, 200), width=line_width)

    # Bottom V fold
    fold_tip = rotate_point(env_center_x, env_center_y + half_h * 0.1,
                           env_center_x, env_center_y, rot_rad)
    draw.line([rotated[3], fold_tip], fill=(200, 200, 200), width=line_width)
    draw.line([rotated[2], fold_tip], fill=(200, 200, 200), width=line_width)

    # Add motion lines (streaks showing movement toward center)
    for i in range(3):
        angle = -30 + i * 15 + rotation
        start_dist = radius * 0.7 + i * 5
        end_dist = radius * 0.5

        start_x = env_center_x + half_w + start_dist * 0.3 * math.cos(math.radians(angle + 180))
        start_y = env_center_y + start_dist * 0.2 * math.sin(math.radians(angle + 180))
        end_x = env_center_x + half_w * 0.8
        end_y = env_center_y - half_h * 0.3 + i * 3

        # Motion streak
        streak_alpha = 120 - i * 30
        draw.line(
            [(start_x, start_y), (end_x, end_y)],
            fill=(255, 255, 255, streak_alpha),
            width=max(1, size // 48 - i)
        )

    # Circular mask
    mask = Image.new('L', (size, size), 0)
    mask_draw = ImageDraw.Draw(mask)
    mask_draw.ellipse(
        [padding, padding, padding + circle_size - 1, padding + circle_size - 1],
        fill=255
    )
    img.putalpha(mask)

    return img


if __name__ == '__main__':
    sizes = [16, 32, 48, 64, 128, 256]
    icons = [create_icon(size) for size in sizes]

    icons[-1].save(output_ico, format='ICO', sizes=[(s, s) for s in sizes], append_images=icons[:-1])
    print(f"Created {output_ico}")

    icons[3].save(output_png, format='PNG')
    print(f"Created {output_png}")

    print("\nDocuShuttle icon:")
    print("  - Tilted envelope being pulled into vortex")
    print("  - Dynamic swirl effect")
    print("  - Motion lines showing movement")
