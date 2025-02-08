import streamlit as st
import json
import win32com.client
import pythoncom


#############################################
# EA COM Connection & Package Collection
#############################################

def get_ea_repository():
    """
    Connect to EA via COM.
    EA must be running with a project open.
    """
    try:
        pythoncom.CoInitialize()  # Initialize COM for this thread
        ea_app = win32com.client.Dispatch("EA.App")
        repository = ea_app.Repository
        return repository
    except Exception as e:
        st.error(f"Could not connect to EA. Is EA running? Error: {e}")
        return None


def recursive_collect_packages(ea_package):
    """
    Recursively traverse the package hierarchy starting at ea_package.
    Returns a list of EA package objects.
    """
    packages = [ea_package]
    sub_packages = ea_package.Packages
    for i in range(sub_packages.Count):
        sub_pkg = sub_packages.GetAt(i)
        packages.extend(recursive_collect_packages(sub_pkg))
    return packages


def get_all_packages(repository):
    """
    Collect all packages from the EA project.
    Typically, repository.Models holds the top-level models.
    """
    packages = []
    models = repository.Models
    for i in range(models.Count):
        model = models.GetAt(i)
        packages.extend(recursive_collect_packages(model))
    return packages


#############################################
# Graph Generation (Mimicking EA Script Structure)
#############################################

def generate_graph_from_package(ea_pkg, repository):
    """
    Generate a JSON graph from the selected EA package.
    The structure (with "nodes" and "edges") is modeled after the EA JScript.
    """
    # Initialize graph and tracking sets/dicts
    graph = {"nodes": [], "edges": []}
    visited_elements = set()
    visited_diagrams = set()
    edge_set = set()
    processed_elements = {}  # Map ElementID -> processed node

    def process_main_diagram(package):
        if package.Diagrams.Count > 0:
            mainDiagram = package.Diagrams.GetAt(0)
            visited_diagrams.add(mainDiagram.DiagramID)
            mainDiagramNode = {
                "id": "D" + str(mainDiagram.DiagramID),
                "name": mainDiagram.Name,
                "type": "Diagram",
                "elements": []
            }
            for e in range(mainDiagram.DiagramObjects.Count):
                diag_object = mainDiagram.DiagramObjects.GetAt(e)
                diag_element = repository.GetElementByID(diag_object.ElementID)
                if diag_element is None:
                    continue
                if diag_element.ElementID not in visited_elements:
                    processedElement = process_element(diag_element, diag_object)
                    if processedElement is not None:
                        mainDiagramNode["elements"].append(processedElement)
            # Add the package node with its main diagram to the graph nodes
            graph["nodes"].append({
                "name": package.Name,
                "type": "Package",
                "diagrams": [mainDiagramNode]
            })

    def process_element(element, diag_object):
        if element.ElementID in visited_elements:
            return processed_elements.get(element.ElementID, None)
        visited_elements.add(element.ElementID)
        node = {
            "id": element.ElementID,
            "name": element.Name,
            "type": element.Type,
            "attributes": [],
            "linkedDiagrams": [],
            "position": {
                "left": getattr(diag_object, "left", None),
                "right": getattr(diag_object, "right", None),
                "top": getattr(diag_object, "top", None),
                "bottom": getattr(diag_object, "bottom", None)
            } if diag_object is not None else None
        }
        # Process element attributes
        for a in range(element.Attributes.Count):
            attribute = element.Attributes.GetAt(a)
            node["attributes"].append({
                "id": attribute.AttributeID,
                "name": attribute.Name,
                "type": attribute.Type,
                "default": attribute.Default
            })
        # Process external classifier if one is referenced
        if element.ClassifierID != 0:
            process_external_classifier(element.ClassifierID)
        # Process any linked diagrams (child diagrams)
        process_linked_diagrams(element, node)
        # Process connectors (edges) for this element
        add_edges_from_element(element)
        # Add this element node to the graph and record it
        graph["nodes"].append(node)
        processed_elements[element.ElementID] = node
        return node

    def process_external_classifier(classifierID):
        classifier = repository.GetElementByID(classifierID)
        if classifier and classifier.ElementID not in visited_elements:
            externalNode = {
                "id": classifier.ElementID,
                "name": classifier.Name,
                "type": classifier.Type,
                "package": None,
                "attributes": []
            }
            pkg = repository.GetPackageByID(classifier.PackageID)
            if pkg:
                externalNode["package"] = pkg.Name
            for a in range(classifier.Attributes.Count):
                attribute = classifier.Attributes.GetAt(a)
                externalNode["attributes"].append({
                    "id": attribute.AttributeID,
                    "name": attribute.Name,
                    "type": attribute.Type,
                    "default": attribute.Default
                })
            graph["nodes"].append(externalNode)
            visited_elements.add(classifier.ElementID)
            processed_elements[classifier.ElementID] = externalNode

    def process_linked_diagrams(element, node):
        for d in range(element.Diagrams.Count):
            diagram = element.Diagrams.GetAt(d)
            if diagram.DiagramID in visited_diagrams:
                continue  # Skip already processed diagrams
            visited_diagrams.add(diagram.DiagramID)
            elementsInDiagram = []
            for e in range(diagram.DiagramObjects.Count):
                diag_object = diagram.DiagramObjects.GetAt(e)
                diag_element = repository.GetElementByID(diag_object.ElementID)
                if diag_element and diag_element.ElementID not in visited_elements:
                    processedElement = process_element(diag_element, diag_object)
                    if processedElement is not None:
                        elementsInDiagram.append(processedElement)
            node["linkedDiagrams"].append({
                "id": "D" + str(diagram.DiagramID),
                "name": diagram.Name,
                "type": "Diagram",
                "elements": elementsInDiagram
            })

    def add_edges_from_element(element):
        for j in range(element.Connectors.Count):
            connector = element.Connectors.GetAt(j)
            sourceElement = repository.GetElementByID(connector.ClientID)
            targetElement = repository.GetElementByID(connector.SupplierID)
            edgeID = f"{sourceElement.ElementID}-{targetElement.ElementID}-{connector.Type}"
            if edgeID not in edge_set:
                edge = {
                    "from": sourceElement.ElementID,
                    "to": targetElement.ElementID,
                    "type": connector.Type,
                    "name": connector.Name
                }
                graph["edges"].append(edge)
                edge_set.add(edgeID)

    # Start processing by handling the package's main diagram
    process_main_diagram(ea_pkg)
    return graph


#############################################
# Streamlit App UI
#############################################

st.title("EA Model Graph Generator")
st.markdown("""
This app connects to Enterprise Architect (EA), retrieves packages from your project,
and generates a JSON graph (nodes & edges) for a selected package following the EA script structure.
""")

# Connect to EA
repository = get_ea_repository()
if repository:
    # Get all packages from EA
    all_packages = get_all_packages(repository)
    # Filter out top-level models (usually ParentID == 0)
    valid_packages = [pkg for pkg in all_packages if pkg.ParentID != 0]
    # Sort packages by name for ease of selection
    valid_packages = sorted(valid_packages, key=lambda p: p.Name)

    if valid_packages:
        # Create a mapping from package names to EA package objects
        package_dict = {pkg.Name: pkg for pkg in valid_packages}
        selected_pkg_name = st.selectbox("Select a package", list(package_dict.keys()))

        if st.button("Generate JSON Graph"):
            selected_pkg = package_dict.get(selected_pkg_name)
            if selected_pkg:
                graph = generate_graph_from_package(selected_pkg, repository)
                st.success("JSON graph created!")
                st.json(graph)
                json_str = json.dumps(graph, indent=4)
                st.download_button("Download JSON", data=json_str, file_name="main_diagram_with_edges.json",
                                   mime="application/json")
            else:
                st.error("Selected package could not be found.")
    else:
        st.error("No valid packages found in the EA project.")
else:
    st.error("Could not connect to EA. Make sure EA is running with a project open.")
