import json

# Paths to your JSON files
chemin = "C:/Users/adel-mounir.achir/OneDrive - GS1/Bureau/Mes documents/Codes and automations/Fiche produit PGC Relooking/"
json_file_path = chemin + "Json Files/"

# Reload the necessary files from the provided paths
with open(json_file_path + "gdsn_classes.json", "r", encoding="utf-8") as f:
    classes_data = json.load(f)
with open(json_file_path + "gdsn_codeValues.json", "r", encoding="utf-8") as f:
    code_lists_data = json.load(f)
with open(json_file_path + "gdsn_instances.json", "r", encoding="utf-8") as f:
    instances = json.load(f)
with open(json_file_path + "gdsn_classAttributes.json", "r", encoding="utf-8") as f:
    attributes_data = json.load(f)

# # Step 1: Create a mapping of attributes to their code values
# attribute_to_codes = {}
# for code in gdsn_code_values:
#     attribute_id = code.get("attributeId")
#     if attribute_id not in attribute_to_codes:
#         attribute_to_codes[attribute_id] = []
#     attribute_to_codes[attribute_id].append({
#         "codeValue": code.get("codeValue"),
#         "definition": code.get("definition")
#     })

# # Step 2: Group attributes under their respective classes
# class_to_attributes = {}
# for attr in gdsn_class_attributes:
#     class_id = attr.get("parentClassId")
#     if class_id not in class_to_attributes:
#         class_to_attributes[class_id] = []
#     class_to_attributes[class_id].append({
#         "attributeId": attr["id"],
#         "name": attr["name"],
#         "definition": attr["definition"],
#         "codes": attribute_to_codes.get(attr["id"], [])
#     })

# # Step 3: Link classes, attributes, and code lists
# linked_structure = []
# for g_class in gdsn_classes:
#     class_id = g_class["id"]
#     linked_structure.append({
#         "classId": class_id,
#         "className": g_class["name"],
#         "definition": g_class["definition"],
#         "attributes": class_to_attributes.get(class_id, [])
#     })

# # Output a sample of the linked structure for inspection
# # print("Linked structure sample:")
# # print(json.dumps(linked_structure, indent=4, ensure_ascii=False))

# output_file_path = json_file_path+"linked_structure.json"

# # Save the linked structure to the specified file
# with open(output_file_path, "w", encoding="utf-8") as f:
#     json.dump(linked_structure, f, indent=4, ensure_ascii=False)

# output_file_path 

import json



# Create mappings for quick access
classes = {}
for cls in classes_data:
    cls_id = cls['id']
    cls['attributes'] = []  # Initialize an empty list of attributes
    classes[cls_id] = cls

attributes = {}
for attr in attributes_data:
    attr_id = attr['id']
    attributes[attr_id] = attr


# Create a mapping from dataClassId/classId to code lists
code_lists = {}
for code in code_lists_data:
    class_id = code['classId']
    if class_id not in code_lists:
        code_lists[class_id] = []
    code_lists[class_id].append(code)

# Link attributes to classes
for attr in attributes_data:
    parent_class_id = attr['parentClassId']
    if parent_class_id in classes:
        classes[parent_class_id]['attributes'].append(attr)
    else:
        print(f"Warning: Parent class ID {parent_class_id} not found for attribute '{attr['name']}'")

# Link code lists to attributes based on dataClassId
for attr in attributes_data:
    data_class_id = attr['dataClassId']
    if data_class_id in code_lists:
        attr['code_list'] = code_lists[data_class_id]
    else:
        attr['code_list'] = []

# ** Export data for "PackagingInformationModule" **

# Find the class "PackagingInformationModule"
target_class_name = "PackagingInformationModule"
target_class = None
output_file_path = json_file_path+"all_data_linked.json"
output_file_path_class = json_file_path+target_class_name+".json"
for cls in classes.values():
    if cls['name'] == target_class_name:
        target_class = cls
        break

if target_class:
    # Export the target class data to JSON
    with open(output_file_path_class, 'w') as f:
        json.dump(target_class, f, indent=4)
    print(f"Exported 'PackagingInformationModule' data to 'PackagingInformationModule.json'.")
else:
    print(f"Class '{target_class_name}' not found.")

# ** Export all classes data to JSON **
with open(output_file_path, 'w') as f:
    json.dump(list(classes.values()), f, indent=4)
print("Exported all classes data to 'AllClassesWithAttributesAndCodeLists.json'.")
