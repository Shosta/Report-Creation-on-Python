'''
Xml file Helpers functions to read and write files, access elements or attributes.
'''

def get_element_value(xml_file, element_tag_name):
    '''Return the value of the first element that has the 'element_tag_name' type item in the
    'xml_file'.'''

    from xml.dom import minidom
    xmldom = minidom.parse(xml_file)
    element = xmldom.getElementsByTagName(element_tag_name)

    first_data = ''
    try:
        first_data = list(element)[0].firstChild.data
    except:
        pass

    return first_data

def get_attribute_value(xml_file, element_tag_name, attribute_name):
    '''Return the value of the first element that has the 'element_tag_name' type item in the
    'xml_file'.'''

    from xml.dom import minidom
    xmldom = minidom.parse(xml_file)
    element = xmldom.getElementsByTagName(element_tag_name)

    attribute_value = ''
    try:
        attribute_value = list(element)[0].getAttribute(attribute_name)
    except IndexError:
        raise Exception('Element ' + element_tag_name + ' does not exist in the document.')

    return attribute_value


def __replaceText(node, new_text):
    if node.firstChild.nodeType != node.TEXT_NODE:
        raise Exception("node does not contain text")

    node.firstChild.replaceWholeText(new_text)


def replace_element_value(xml_dom, element_tag_name, value):
    try:
        node = xml_dom.getElementsByTagName(element_tag_name)[0]
        __replaceText(node, value)
    except IndexError:
        raise Exception('Element ' + element_tag_name + ' does not exist in the document.')
    #print(xmldom.toxml())


def replace_attribute_value(xml_dom, element_tag_name, attribute_name, value):
    try:
        node = xml_dom.getElementsByTagName(element_tag_name)[0]
        node.setAttribute(attribute_name, value)
    except IndexError:
        raise Exception('Element ' + element_tag_name + ' does not exist in the document.')


