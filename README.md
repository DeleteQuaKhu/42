from sklearn.cluster import DBSCAN
from pptx import Presentation
from pptx.util import Inches, Pt
from PIL import Image
import os
import re
import glob
import numpy as np

def generate_modified_lists(point,tolorance):
    return [point[0]+tolorance,point[0]-tolorance,point[1]+tolorance,point[1]-tolorance,point[2]+tolorance,point[2]-tolorance]
def center(list_node):
    center_coordinate = [0,0,0] 
    for i in range(0,len(list_node)):
        center_coordinate[0] += float(list_node[i][0])/len(list_node)
        center_coordinate[1] += float(list_node[i][1])/len(list_node)
        center_coordinate[2] += float(list_node[i][2])/len(list_node)
    return center_coordinate
def cluster_points(data, max_distance):
    coordinates = np.array([item[1:] for item in data])
    clustering = DBSCAN(eps=max_distance, min_samples=1).fit(coordinates)
    clusters = {}
    for point, label in zip(data, clustering.labels_):
        clusters.setdefault(label, []).append(point)
    return list(clusters.values())


def insert_images_at_text_positions(ppt_path, img_directory):
    prs = Presentation(ppt_path)
    number_pattern = re.compile(r'^\d+$')  # Regular expression to match standalone numbers

    for slide in prs.slides:
        for shape in list(slide.shapes):  # Create a static list of shapes
            if shape.has_text_frame and shape.text_frame.text:
                text = shape.text_frame.text.strip()
                match = number_pattern.fullmatch(text)
                if match:
                    number = int(match.group())
                    # Find an image file that starts with "image_{number}_" and ends with ".png"
                    img_files = glob.glob(os.path.join(img_directory, f"image_{number}_*.png"))
                    if img_files:
                        # If there's more than one match, just take the first one
                        img_path = img_files[0]

                        # Get the position of the text box
                        left = shape.left
                        top = shape.top

                        # Load the image to get its dimensions
                        with Image.open(img_path) as img:
                            # Convert pixel to EMUs
                            width = Inches(img.width / img.info['dpi'][0])
                            height = Inches(img.height / img.info['dpi'][1])

                        # Add the image to the slide at the position of the text box
                        slide.shapes.add_picture(img_path, left, top, width, height)

                        # Delete the text box
                        sp = shape._element
                        sp.getparent().remove(sp)

    # Save the modified presentation
    prs.save('modified_presentation.pptx')

# Example usage
ppt_path = r"C:\Users\TechnoStar\Documents\Valve_missalignment\ver2\42\New Microsoft PowerPoint Presentation.pptx"
img_directory = r"C:\Users\TechnoStar\Documents\Valve_missalignment\ver2\42"
insert_images_at_text_positions(ppt_path, img_directory)
##########################################################
###########   transform 3/3/2024  ########################
##########################################################
import numpy as np

def distance(point1, point2):
    # Tính khoảng cách giữa hai điểm trong không gian 3D
    return np.sqrt((point1[0]-point2[0])**2 + (point1[1]-point2[1])**2 + (point1[2]-point2[2])**2)

def transform_points(big_list):
    first_point = np.array(big_list[0][1:])
    last_point = np.array(big_list[-1][1:])
    line_vector = last_point - first_point
    
    transformed_points = []
    for item in big_list:
        point = np.array(item[1:])
        projected_point = first_point + np.dot(point - first_point, line_vector) / np.dot(line_vector, line_vector) * line_vector
        transformed_points.append(projected_point)
    
    return transformed_points

def sort_by_distance(transformed_points):
    # Tính toán khoảng cách từ các điểm mới đến điểm đầu
    distances = [distance(np.array([0, 0, 0]), point) for point in transformed_points]
    
    # Sắp xếp lại big_list theo khoảng cách tăng dần
    sorted_indices = np.argsort(distances)
    sorted_points = [transformed_points[i] for i in sorted_indices]
    
    return sorted_points

def transform_and_sort(big_list):
    transformed_points = transform_points(big_list)
    sorted_points = sort_by_distance(transformed_points)
    
    return sorted_points

# Example usage
big_list = [[1232, 0, 1, 3], [112, 0, 1, 0], [1332, 0, 2, 3]]
sorted_points = transform_and_sort(big_list)
print(sorted_points)

big_list = [[132, 1, 3, 5], [122, 2, 4, 12], [131, 3, 5, 7], [221, 4, 6, 4]]

print(sorted(big_list, key=lambda z: z[3]))


