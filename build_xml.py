from xml.etree.ElementTree import Element, SubElement, tostring
import xml.dom.minidom

dct_testcase = {
    "summary": "<![CDATA[Verify installation with default settings completes correctly]]>",
    "precondition": "<![CDATA[]]>",
    "execution_type": "<![CDATA[]]>",
    "importance": "<![CDATA[]]>"}


def dict_to_elem(dictionary):
    item = Element('Item')
    for key in dictionary:
        field = Element(key.replace(' ', ''))
        field.text = dictionary[key]
        item.append(field)
    return itemroot = Element('testsuite', name='')


parent = SubElement(root, 'testsuite', name='Installation')

child_details = SubElement(parent, 'details')
child_details.text = '<![CDATA[<p></p>]]>'

child_testcase = SubElement(parent, 'testcase', name='Normal installation')
child_testcase.text = '<![CDATA[Verify installation with default settings completes correctly]]>'

child_testcase.append(dict_to_elem(dct_testcase))

child_with_entity_ref = SubElement(root, 'child_with_entity_ref')
child_with_entity_ref.text = 'This & that'

xml_string = tostring(root)

xml = xml.dom.minidom.parseString(xml_string)
pretty_xml_as_string = xml.toprettyxml()
print(pretty_xml_as_string)
