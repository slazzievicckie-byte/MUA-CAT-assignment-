from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH


def add_paragraph(doc, text, bold=False, center=False, size=12):
    p = doc.add_paragraph()
    if center:
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(text)
    run.font.name = 'Times New Roman'
    run.font.size = Pt(size)
    if bold:
        run.bold = True
    return p


def add_cover_page(doc):
    add_paragraph(doc, 'MANAGEMENT UNIVERSITY OF AFRICA', bold=True, center=True, size=18)
    add_paragraph(doc, 'BDS 106 - SOCIOLOGY OF DEVELOPMENT (CAT)', bold=True, center=True, size=18)
    doc.add_paragraph()
    add_paragraph(doc, 'INSTRUCTIONS:', bold=True, size=12)

    instructions = [
        '1. Use Times New Roman font type.',
        '2. Font size 12.',
        '3. Clearly show the beginning of every new concept by using paragraphs.',
        '4. Have a cover page that shows Name of University, student full name, student Admission number, Unit title, Unit code, nature of assignment (CAT or WBA), Lecturer name, Semester and Year.',
        '5. Use APA style – Reference & Citation.',
        '6. Use clear examples.',
        '7. Use ideas from the notes presented and enhance your assignment by undertaking comprehensive research using relevant library materials.',
        '8. All students must hand in a soft copy through the Moodle online system.',
        '9. No assignment will be accepted after the due date.'
    ]
    for item in instructions:
        add_paragraph(doc, item, size=12)

    doc.add_page_break()

    add_paragraph(doc, 'SAMPLE COVER PAGE', bold=True, center=True, size=16)
    add_paragraph(doc, 'MANAGEMENT UNIVERSITY OF AFRICA', bold=True, center=True, size=14)
    doc.add_paragraph()
    add_paragraph(doc, "STUDENT'S NAME:", bold=True, center=True, size=14)
    add_paragraph(doc, 'ADMISSION NUMBER: BDS1/00900/3/2011', bold=True, center=True, size=14)
    add_paragraph(doc, 'UNIT TITLE: SOCIOLOGY OF DEVELOPMENT', bold=True, center=True, size=14)
    add_paragraph(doc, 'UNIT CODE: BDS 106', bold=True, center=True, size=14)
    add_paragraph(doc, 'WORK BASED ASSIGNMENT', bold=True, center=True, size=14)
    add_paragraph(doc, "LECTURER'S NAME: EDNA MORIASI", bold=True, center=True, size=14)
    add_paragraph(doc, 'SEMESTER: MAY – AUGUST 2026', bold=True, center=True, size=14)

    doc.add_page_break()


def add_question(doc, question, answer):
    add_paragraph(doc, question, bold=True, size=12)
    paragraphs = answer.split('\n\n')
    for paragraph in paragraphs:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        run = p.add_run(paragraph)
        run.font.name = 'Times New Roman'
        run.font.size = Pt(12)
    doc.add_paragraph()


def build_docx():
    doc = Document()
    add_cover_page(doc)
    add_paragraph(doc, 'CONTINUOUS ASSESSMENT TEST', bold=True, center=True, size=16)
    add_paragraph(doc, 'MAXIMUM MARKS: 15', size=12)
    add_paragraph(doc, 'INSTRUCTIONS:', bold=True, size=12)
    add_paragraph(doc, 'ANSWER ALL QUESTIONS', bold=True, size=12)
    doc.add_paragraph()

    q1 = 'a) Assess five characteristics of sociology and explain why they make it a distinct field of study. (5 Marks)'
    a1 = (
        'Sociology is a distinct field of study because it is concerned with the scientific study of society, social relationships and social behaviour. It differs from other disciplines because it focuses on social life as a whole, not only on the individual. The following are five major characteristics of sociology (Giddens, 2009; Macionis, 2017).\n\n'
        'First, sociology is a scientific and systematic discipline. It collects evidence through observation, interviews, surveys and analysis before making conclusions. This makes sociology different from personal opinion or speculation. According to Giddens (2009), sociology uses systematic methods to investigate social reality in a disciplined and evidence-based manner.\n\n'
        'Second, sociology studies society as a whole. It examines how people live together in families, communities, schools, workplaces and nations. It focuses on how human beings create and maintain social order. Social life cannot be understood only by focusing on one individual because people are shaped by the environment, institutions and social groups around them (Macionis, 2017).\n\n'
        'Third, sociology studies social institutions. It analyses how institutions such as the family, education, religion, economy and government shape behaviour and values. These institutions influence how people think, act and relate to one another. This demonstrates that human behaviour is not only personal but also socially constructed and structured (Kivisto, 2013).\n\n'
        'Fourth, sociology is comparative and broad in scope. It compares societies across time and space, showing how cultures, values and social structures differ. This helps us understand how societies change and develop. A sociologist is interested in patterns that exist across societies rather than only the experience of one community (Giddens, 2009).\n\n'
        'Fifth, sociology is concerned with social problems and social change. It studies issues such as poverty, crime, unemployment, inequality and migration. It seeks to explain their causes and how they affect society. In this way, sociology does not only describe society but also attempts to explain social transformation and social challenges (Macionis, 2017).\n\n'
        'These characteristics make sociology a distinct field because it goes beyond individual behaviour and examines the patterns, systems and institutions that shape social life. It is different from psychology, which focuses mainly on the individual, and economics, which concentrates mainly on production and wealth. Sociology therefore provides a broader understanding of how society functions and changes (Giddens, 2009; Macionis, 2017).\n\n'
        'Reference: Giddens, A. (2009). Sociology (6th ed.). Polity Press. Macionis, J. J. (2017). Sociology (15th ed.). Pearson. Kivisto, P. (2013). Social Theory: Roots and Branches (4th ed.). Oxford University Press.'
    )
    add_question(doc, q1, a1)

    q2 = 'b) Discuss the importance of sociology in understanding contemporary Kenyan society. (5 Marks)'
    a2 = (
        'Sociology is very important in understanding contemporary Kenyan society because Kenya is undergoing rapid social, economic and political transformation. It helps people explain and analyse the changes occurring in the country. A social understanding of Kenya is essential because development is shaped not only by economic policy but also by social relations, culture and institutions (Macionis, 2017).\n\n'
        'First, sociology helps explain social problems in Kenya. Issues such as unemployment, poverty, inequality, youth unrest, crime and corruption are not isolated events; they are influenced by social policies, economic structures and cultural systems. Sociology helps identify their causes and possible solutions. This is important because social problems are often rooted in structural inequality and unequal access to resources (Giddens, 2009).\n\n'
        'Second, sociology helps us understand social inequality. In Kenya, access to resources, opportunities and services is often influenced by class, gender, ethnicity, religion and regional location. Sociology helps reveal how these factors affect education, employment and social mobility. Such inequalities shape social life and limit the ability of many people to participate fully in national development (Kivisto, 2013).\n\n'
        'Third, sociology helps explain culture and identity. Kenya is a multicultural society with diverse ethnic groups, languages and traditions. Sociology helps us understand how culture shapes values, beliefs and social behaviour, while also promoting national unity and social cohesion. Understanding cultural diversity is crucial in a multi-ethnic society such as Kenya (Macionis, 2017).\n\n'
        'Fourth, sociology is useful in understanding political and institutional changes. Kenya has experienced reforms in governance, education, devolution and public participation. Sociology helps explain how institutions affect social life and how people respond to political and administrative changes. It also highlights how social trust, civic engagement and institutional legitimacy influence social order (Giddens, 2009).\n\n'
        'Fifth, sociology contributes to social planning and development. Governments, NGOs and development agencies need sociological insight to design effective policies in education, health, housing, employment and community development. This makes sociology relevant to planning for national development. Development policies are more effective when they respond to real social conditions rather than assumptions (World Bank, 2022).\n\n'
        'In conclusion, sociology is important in understanding Kenyan society because it enables people to analyse social reality, explain social problems and support better development and policy-making. It helps create a more informed and just society. Sociology therefore remains essential for understanding the challenges and opportunities facing contemporary Kenya (World Bank, 2022; Kenya National Bureau of Statistics, 2023).\n\n'
        'Reference: Macionis, J. J. (2017). Sociology (15th ed.). Pearson. Giddens, A. (2009). Sociology (6th ed.). Polity Press. Kivisto, P. (2013). Social Theory: Roots and Branches (4th ed.). Oxford University Press. Kenya National Bureau of Statistics. (2023). Statistical Abstract. Nairobi. World Bank. (2022). Kenya Economic Update. Washington, DC.'
    )
    add_question(doc, q2, a2)

    q3 = 'c) Discuss three major theories of social change, and explain how they can be applied to Kenya’s ongoing development. (5 Marks)'
    a3 = (
        'Three major theories of social change are evolutionary theory, conflict theory and functionalist theory. These theories explain how societies change and how development occurs. Each theory offers a different explanation of social transformation (Giddens, 2009; Macionis, 2017).\n\n'
        'Evolutionary theory argues that societies change gradually from simple forms to more complex forms. It sees development as a process of progress, improvement and modernization. In Kenya, this theory can be applied to the country’s growth in education, urbanization, technology and infrastructure. For example, the expansion of mobile technology, internet use, higher education and modern transport are evidence of social evolution. This shows how societies progress as they adopt new ideas, technologies and institutions (Kivisto, 2013).\n\n'
        'Conflict theory explains change as the result of conflict between groups with competing interests. It argues that social change occurs when underprivileged or marginalized groups challenge existing inequalities. In Kenya, this theory can be applied to conflicts over land, unemployment, unequal distribution of resources and political competition. These conflicts often lead to reforms and social change. In many societies, power struggles and deprivation are central to social transformation (Macionis, 2017).\n\n'
        'Functionalist theory views society as an interrelated system where each part contributes to stability and order. Social change happens when institutions adjust to new needs. In Kenya, this theory is relevant in understanding reforms in education, health, devolution and governance. When institutions adapt to social demands, society becomes more stable and better able to develop. This approach highlights the need for institutional consistency and social integration during periods of change (Giddens, 2009).\n\n'
        'These theories are relevant to Kenya’s development because the country is undergoing modernization, social inequality, political reform and institutional change. Evolutionary theory explains gradual development, conflict theory explains struggles over resources and power, and functionalist theory explains continuity and adaptation in social systems. Together, they provide a complete understanding of Kenya’s social change and development process (Giddens, 2009; Macionis, 2017; Kivisto, 2013).\n\n'
        'Reference: Giddens, A. (2009). Sociology (6th ed.). Polity Press. Macionis, J. J. (2017). Sociology (15th ed.). Pearson. Kivisto, P. (2013). Social Theory: Roots and Branches (4th ed.). Oxford University Press. World Bank. (2022). Kenya Economic Update. Washington, DC.'
    )
    add_question(doc, q3, a3)

    # Final references section
    add_paragraph(doc, 'REFERENCES', bold=True, center=True, size=12)
    refs = [
        'Giddens, A. (2009). Sociology (6th ed.). Polity Press.',
        'Macionis, J. J. (2017). Sociology (15th ed.). Pearson.',
        'Kivisto, P. (2013). Social Theory: Roots and Branches (4th ed.). Oxford University Press.',
        'Kenya National Bureau of Statistics. (2023). Statistical Abstract. Nairobi.',
        'World Bank. (2022). Kenya Economic Update. Washington, DC.'
    ]
    for ref in refs:
        add_paragraph(doc, ref, size=11)

    return doc


def main():
    doc = build_docx()
    filename = 'BDS_106_Sociology_of_Development.docx'
    doc.save(filename)
    print(f'Created {filename}')


if __name__ == '__main__':
    main()
