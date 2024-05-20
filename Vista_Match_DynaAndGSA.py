import pandas as pd
from scipy.spatial import KDTree
import matplotlib.pyplot as plt
import numpy as np

list_GSA = r'C:\Users\Elian.Zhang\Arup\PROJECT VISTA - 4 Internal Data\04 Calculations\07 Structural\09 Substructure\Stage 4\12 Basement Design\01_GSA\04_OUTPUTS\240516 Cores base shrinkage forces to GEO\GSA_Shrinkage_nodes.txt'
list_Dyna = r'C:\Users\Elian.Zhang\Arup\PROJECT VISTA - 4 Internal Data\04 Calculations\07 Structural\09 Substructure\Stage 4\12 Basement Design\01_GSA\04_OUTPUTS\240516 Cores base shrinkage forces to GEO\DynaNodes.txt'
output_file = r'C:\Users\Elian.Zhang\Arup\PROJECT VISTA - 4 Internal Data\04 Calculations\07 Structural\09 Substructure\Stage 4\12 Basement Design\01_GSA\04_OUTPUTS\240516 Cores base shrinkage forces to GEO\match.txt'
output_folder = r'C:\Users\Elian.Zhang\Arup\PROJECT VISTA - 4 Internal Data\04 Calculations\07 Structural\09 Substructure\Stage 4\12 Basement Design\01_GSA\04_OUTPUTS\240516 Cores base shrinkage forces to GEO'

# Load data from the first text file (List A)
df_a = pd.read_csv(list_GSA, delimiter='\t')  # Adjust delimiter if needed
points_a = df_a.iloc[:, 1:3].values  # Assuming X is column 2 and Y is column 3
node_ids_a = df_a.iloc[:, 0].values  # Assuming node ID is column 1

# Load data from the second text file (List B)
df_b = pd.read_csv(list_Dyna, delimiter='\t')  # Adjust delimiter if needed
points_b = df_b.iloc[:, 1:3].values  # Assuming X is column 2 and Y is column 3
node_ids_b = df_b.iloc[:, 0].values  # Assuming node ID is column 1

# Build a KDTree for List B
tree = KDTree(points_b)

# Find the closest point in List B for each point in List A
distances, indices = tree.query(points_a)


# Mask to keep track of points in set B that have been assigned
assigned_mask = np.zeros(len(points_b), dtype=bool)

# Create lists to store matched points
matched_points_a = []
matched_points_b = []

# Iterate over points in set A
for i, (point_a, node_id_a) in enumerate(zip(points_a, node_ids_a)):
    # Find the closest unassigned point in set B
    closest_distance = float('inf')
    closest_index = None
    for j, (point_b, node_id_b) in enumerate(zip(points_b, node_ids_b)):
        if not assigned_mask[j]:
            distance = np.linalg.norm(point_a - point_b)
            if distance < closest_distance:
                closest_distance = distance
                closest_index = j
    # Mark the closest point in set B as assigned
    assigned_mask[closest_index] = True
    # Add the matched points to the lists
    matched_points_a.append([node_id_a, point_a[0], point_a[1]])
    matched_points_b.append([node_ids_b[closest_index], points_b[closest_index, 0], points_b[closest_index, 1]])

# Create a DataFrame for the output
output_df = pd.DataFrame({
    'node_ID_A': [point[0] for point in matched_points_a],
    'X_A': [point[1] for point in matched_points_a],
    'Y_A': [point[2] for point in matched_points_a],
    'node_ID_B': [point[0] for point in matched_points_b],
    'X_B': [point[1] for point in matched_points_b],
    'Y_B': [point[2] for point in matched_points_b],
    'distance': [np.linalg.norm(np.array(point_a[1:]) - np.array(point_b[1:])) for point_a, point_b in zip(matched_points_a, matched_points_b)]
})

# Write the output to a text file
output_df.to_csv(output_file, sep='\t', index=False)  # Adjust separator if needed


################
# Plot the points and lines
plt.figure(figsize=(30, 24))

# Plot points from set A
plt.scatter(points_a[:, 0], points_a[:, 1], c='blue', label='GSA nodes')

# Annotate each point in set A with its node ID
for i, txt in enumerate(node_ids_a):
    plt.annotate(txt, (points_a[i, 0], points_a[i, 1]), size=3)

# Plot points from set B closest to points in set A
plt.scatter([point[1] for point in matched_points_b], [point[2] for point in matched_points_b], c='red', label='Dyna nodes closest to GSA nodes')

# Plot lines linking each pair of points
for point_a, point_b in zip(matched_points_a, matched_points_b):
    plt.plot([point_a[1], point_b[1]], [point_a[2], point_b[2]], color='green')

# Set labels and legend
plt.axis('equal')
plt.xlabel('X')
plt.ylabel('Y')
plt.legend()
plt.title('Matched Points and Lines')
plt.grid(True)

plt.savefig(output_folder + "\plot.png", dpi=300)

plt.close()

########################

# Plot the points from both sets A and B
plt.figure(figsize=(100, 80))

# Plot points from set A
plt.scatter(points_a[:, 0], points_a[:, 1], c='blue', label='GSA nodes')

# Annotate each point in set A with its node ID
for i, txt in enumerate(node_ids_a):
    plt.annotate(txt, (points_a[i, 0], points_a[i, 1]), size=12)

# Plot points from set B
plt.scatter(points_b[:, 0], points_b[:, 1], c='red', label='Dyna nodes')

# Annotate each point in set B with its node ID
for i, txt in enumerate(node_ids_b):
    plt.annotate(txt, (points_b[i, 0], points_b[i, 1]), size=12)

# Set labels and legend
plt.axis('equal')
plt.xlabel('X')
plt.ylabel('Y')
plt.legend()
plt.title('All Points from Sets A and B')
plt.grid(True)
plt.savefig(output_folder + "\plot_all_points.png", dpi=300)
plt.close()