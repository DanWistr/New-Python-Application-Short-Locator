from PIL import Image
import numpy as np
import os

def downscale_and_bitwise_with_hex(image_path, output_size):
    """
    Downscale an image, compute the bitwise map (inverted colors), 
    and return the hexadecimal representation of the downscaled image.

    Parameters:
        image_path (str): Path to the input image.
        output_size (tuple): Tuple containing the size to downscale to (width, height).

    Returns:
        tuple: 
            - A NumPy array containing the bitwise map of the downscaled image.
            - A string of hexadecimal representation of the downscaled image's binary data.
    """
    try:
        # Open the image
        img = Image.open(image_path)
        
        # Ensure the image is in RGB mode to allow bitwise operations on all 3 channels
        if img.mode != "RGB":
            img = img.convert("RGB")
        
        # Downscale the image to the desired size (uses high-quality LANCZOS resampling)
        img_resized = img.resize(output_size, Image.Resampling.LANCZOS)
        
        # Save intermediate downscaled image to a temporary file
        downscaled_temp_path = "downscaled_temp.jpg"
        img_resized.save(downscaled_temp_path)
        
        # Convert the resized image to a NumPy array (preserves 3 color channels)
        img_array = np.array(img_resized)
        
        # Compute the bitwise NOT operation on all three channels (color inversion)
        bitwise_map = np.bitwise_not(img_array)
        
        # Read the downscaled image's binary data and convert it into hexadecimal
        with open(downscaled_temp_path, "rb") as temp_file:
            binary_data = temp_file.read()
            hex_representation = binary_data.hex()

        # Cleanup the temporary file
        os.remove(downscaled_temp_path)
        
        # Return both the bitwise map as an array and the hex representation
        return bitwise_map, hex_representation
    
    except Exception as e:
        print("Error:", str(e))
        return None, None

# Example usage
if __name__ == "__main__":
    # Path to the input image
    image_path = r"C:\ShortCam II\Record\Pic_250621_132307.JPG"  # Replace with your image path
    
    # Desired downscale size (e.g., 720x720)
    downscale_size = (720, 720)

    # Call the function to process the image
    bitwise_map, hex_data = downscale_and_bitwise_with_hex(image_path, downscale_size)
    
    if bitwise_map is not None and hex_data is not None:
        print("Bitwise map of the image (downscaled to {}):".format(downscale_size))
        
        # Save the final result (bitwise map) as an image for visualization
        bitwise_image = Image.fromarray(bitwise_map)
        bitwise_image.save("bitwise_map.jpg")
        print("Bitwise map saved as 'bitwise_map.jpg'")
        
        # Save the hex data or output part of it for review
        print("Hexadecimal Representation of the Downscaled Image")
        print(hex_data)  # To avoid printing the entire hex string (truncate for display)
        print("\nTotal Length of Hex Data:", len(hex_data))
