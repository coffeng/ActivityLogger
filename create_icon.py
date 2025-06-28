"""
Create a high-quality icon file for ActivityLogger
"""
import math
from PIL import Image, ImageDraw, ImageFilter

def create_stopwatch_icon():
    """Create a high-quality stopwatch icon with anti-aliasing"""
    # Create multiple sizes for proper .ico file with high quality
    sizes = [16, 20, 24, 32, 40, 48, 64, 96, 128, 256]
    images = []
    
    for size in sizes:
        # Create image at 4x size for better anti-aliasing, then resize down
        render_size = size * 4
        img = Image.new('RGBA', (render_size, render_size), (0, 0, 0, 0))
        d = ImageDraw.Draw(img)
        
        # Scale factors based on render size
        scale = render_size / 256.0
        
        # Colors
        watch_body_color = (245, 245, 245, 255)  # Light gray/silver
        watch_border_color = (60, 60, 60, 255)   # Dark gray
        button_color = (40, 40, 40, 255)         # Dark button
        hour_hand_color = (220, 20, 20, 255)     # Red hour hand
        minute_hand_color = (20, 60, 200, 255)   # Blue minute hand
        tick_major_color = (0, 0, 0, 255)        # Black major ticks
        tick_minor_color = (120, 120, 120, 255)  # Gray minor ticks
        
        # Watch body dimensions
        margin = int(8 * scale)
        button_width = int(20 * scale)
        button_height = int(16 * scale)
        button_x = (render_size - button_width) // 2
        button_y = int(4 * scale)
        
        # Main watch circle
        circle_coords = [margin, margin + button_height//2, 
                        render_size - margin, render_size - margin + button_height//2]
        
        # Draw outer shadow for depth
        shadow_offset = int(2 * scale)
        shadow_coords = [coord + shadow_offset for coord in circle_coords]
        d.ellipse(shadow_coords, fill=(0, 0, 0, 40))
        
        # Draw main watch body
        d.ellipse(circle_coords, outline=watch_border_color, 
                 width=max(2, int(4 * scale)), fill=watch_body_color)
        
        # Draw inner rim
        inner_margin = margin + int(8 * scale)
        inner_coords = [inner_margin, inner_margin + button_height//2,
                       render_size - inner_margin, render_size - inner_margin + button_height//2]
        d.ellipse(inner_coords, outline=(180, 180, 180), width=max(1, int(2 * scale)))
        
        # Stopwatch button (crown)
        d.rounded_rectangle([button_x, button_y, button_x + button_width, button_y + button_height], 
                           radius=int(3 * scale), fill=button_color)
        
        # Button highlight
        d.rounded_rectangle([button_x + int(2 * scale), button_y + int(2 * scale), 
                           button_x + button_width - int(2 * scale), button_y + int(4 * scale)], 
                           radius=int(1 * scale), fill=(100, 100, 100))
        
        # Calculate center and radii
        center_x = render_size // 2
        center_y = (render_size + button_height//2) // 2
        radius_outer = (render_size // 2) - margin - int(12 * scale)
        radius_tick_outer = radius_outer - int(4 * scale)
        radius_tick_inner = radius_outer - int(12 * scale)
        radius_minor_outer = radius_outer - int(2 * scale)
        radius_minor_inner = radius_outer - int(6 * scale)
        
        # Draw major hour ticks (12, 3, 6, 9)
        for angle in [0, 90, 180, 270]:
            rad = math.radians(angle - 90)  # Start from 12 o'clock
            x1 = center_x + radius_tick_inner * math.cos(rad)
            y1 = center_y + radius_tick_inner * math.sin(rad)
            x2 = center_x + radius_tick_outer * math.cos(rad)
            y2 = center_y + radius_tick_outer * math.sin(rad)
            d.line([x1, y1, x2, y2], fill=tick_major_color, width=max(2, int(4 * scale)))
        
        # Draw minor ticks
        for angle in [30, 60, 120, 150, 210, 240, 300, 330]:
            rad = math.radians(angle - 90)
            x1 = center_x + radius_minor_inner * math.cos(rad)
            y1 = center_y + radius_minor_inner * math.sin(rad)
            x2 = center_x + radius_minor_outer * math.cos(rad)
            y2 = center_y + radius_minor_outer * math.sin(rad)
            d.line([x1, y1, x2, y2], fill=tick_minor_color, width=max(1, int(2 * scale)))
        
        # Watch hands (classic 10:10 position)
        hand_radius_hour = radius_tick_inner * 0.55
        hand_radius_minute = radius_tick_inner * 0.75
        
        # Hour hand (pointing to 10) - 300 degrees
        hour_angle = math.radians(300 - 90)
        hour_x = center_x + hand_radius_hour * math.cos(hour_angle)
        hour_y = center_y + hand_radius_hour * math.sin(hour_angle)
        
        # Draw hour hand with rounded end
        hand_width = max(3, int(6 * scale))
        d.line([center_x, center_y, hour_x, hour_y], fill=hour_hand_color, width=hand_width)
        d.ellipse([hour_x - hand_width//2, hour_y - hand_width//2, 
                  hour_x + hand_width//2, hour_y + hand_width//2], fill=hour_hand_color)
        
        # Minute hand (pointing to 2) - 60 degrees
        minute_angle = math.radians(60 - 90)
        minute_x = center_x + hand_radius_minute * math.cos(minute_angle)
        minute_y = center_y + hand_radius_minute * math.sin(minute_angle)
        
        # Draw minute hand with rounded end
        minute_hand_width = max(2, int(4 * scale))
        d.line([center_x, center_y, minute_x, minute_y], fill=minute_hand_color, width=minute_hand_width)
        d.ellipse([minute_x - minute_hand_width//2, minute_y - minute_hand_width//2,
                  minute_x + minute_hand_width//2, minute_y + minute_hand_width//2], fill=minute_hand_color)
        
        # Center hub
        hub_size = max(4, int(8 * scale))
        d.ellipse([center_x - hub_size, center_y - hub_size, 
                  center_x + hub_size, center_y + hub_size], 
                 fill=(40, 40, 40), outline=(0, 0, 0), width=max(1, int(2 * scale)))
        
        # Center dot highlight
        highlight_size = max(2, int(4 * scale))
        d.ellipse([center_x - highlight_size, center_y - highlight_size,
                  center_x + highlight_size, center_y + highlight_size], fill=(200, 200, 200))
        
        # Resize down with high quality anti-aliasing
        img = img.resize((size, size), Image.LANCZOS)
        
        # Apply slight sharpening for small sizes
        if size <= 32:
            img = img.filter(ImageFilter.UnsharpMask(radius=0.5, percent=50, threshold=0))
        
        images.append(img)
    
    return images

def main():
    """Create and save the high-quality icon file"""
    print("Creating high-quality stopwatch icon...")
    
    try:
        images = create_stopwatch_icon()
        
        if images:
            # Save as .ico file with all sizes
            images[0].save('icon.ico', format='ICO', sizes=[(img.width, img.height) for img in images])
            print("High-quality icon saved as icon.ico")
            
            # Save largest size as PNG for preview
            images[-1].save('icon_preview.png', format='PNG')
            print("Preview saved as icon_preview.png")
            
            print(f"Created icon with {len(images)} sizes: {[img.size for img in images]}")
        else:
            print("Error: No images created")
            
    except ImportError:
        print("Error: PIL (Pillow) not installed. Run: pip install pillow")
    except Exception as e:
        print(f"Error creating icon: {e}")

if __name__ == "__main__":
    main()