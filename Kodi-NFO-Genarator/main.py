import os
import xml.etree.ElementTree as ET

path = "/root/Downloads/Android/"
types = ("mp4", "mkv")


def main():
    load()


def load():
    files = os.listdir(path)

    for f in files:
        if f.endswith(types):
            title = videotitle(f)
            ftitle = nfo_title(f)
            create_nfo_file(ftitle, title)


def videotitle(file):
    title = (str(os.path.splitext(file)[0]).split("-")[2:][0].strip())
    return (title)


def nfo_title(file):
    nfotitle = (os.path.splitext(file)[0] + ".nfo")
    return nfotitle


def create_nfo_file(ftitle, vtitle):
    print("[+] Creating NFO-File")

    # create elements in tree
    parent = ET.Element("episodedetails")
    title = ET.SubElement(parent, "title")
    unic_id = ET.SubElement(parent, "id")
    plot = ET.SubElement(parent, "plot")
    aired = ET.SubElement(parent, "aired")
    runtime = ET.SubElement(parent, "runtime")
    credit = ET.SubElement(parent, "credits")

    # fill elements
    title.text = vtitle
    plot.text = "plot text"
    unic_id.text = "Android"
    aired.text = "2022-11-20"
    runtime.text = "5"
    credit.text = " "

    # create indent and write xml file
    tree = ET.ElementTree(parent)
    ET.indent(tree, space="\t", level=0)
    tree.write(
        (f"{path}{ftitle}"),
        xml_declaration=True,
        encoding='UTF-8',
        method='xml'
    )


if __name__ == "__main__":
    main()
