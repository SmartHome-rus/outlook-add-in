# Assets Folder

This folder contains the icon files referenced in the Outlook add-in manifest.

## Icon Files

- **icon-16.svg** - 16x16 pixel icon for small UI elements
- **icon-32.svg** - 32x32 pixel icon for standard display
- **icon-64.svg** - 64x64 pixel icon for high-resolution displays  
- **icon-80.svg** - 80x80 pixel icon for high-DPI scenarios

## Current Icons

The current icons are placeholder SVG files featuring:
- Blue background (#0078d4) representing the Microsoft Office color scheme
- White "T" letter representing "Test" functionality
- Orange warning dot (#ff6b35) on larger icons to indicate the checking functionality

## Replacing Icons

To use custom icons for your organization:

1. **Create PNG or SVG files** with the exact dimensions listed above
2. **Replace the existing files** with your custom icons
3. **Update the manifest.xml** if changing file extensions:
   ```xml
   <!-- For PNG files -->
   <IconUrl DefaultValue="https://your-domain.com/assets/icon-32.png" />
   
   <!-- For SVG files -->
   <IconUrl DefaultValue="https://your-domain.com/assets/icon-32.svg" />
   ```

## Design Guidelines

Follow Microsoft's Office Add-in icon design guidelines:

- **Use simple, recognizable symbols**
- **Maintain good contrast** for visibility
- **Keep consistent styling** across all sizes
- **Test on different backgrounds** (light/dark themes)
- **Ensure scalability** especially for SVG files

## File Format Recommendations

- **SVG**: Recommended for crisp display at all sizes and resolutions
- **PNG**: Alternative option, ensure high DPI versions for retina displays
- **Avoid**: JPEG (poor quality for icons), GIF (limited colors)

## Technical Requirements

- Icons must be accessible via HTTPS when deployed
- File size should be minimized for faster loading
- SVG files should not contain external dependencies or scripts
- PNG files should use transparency where appropriate