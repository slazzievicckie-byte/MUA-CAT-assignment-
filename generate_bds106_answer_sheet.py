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
    add_paragraph(doc, 'MANAGEMENT UNIVERSITY OF AFRICA', bold=True, center=True, size=20)
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


def build_answer_sheet():
    doc = Document()
    add_cover_page(doc)

    add_paragraph(doc, 'CONTINUOUS ASSESSMENT TEST', bold=True, center=True, size=16)
    add_paragraph(doc, 'MAXIMUM MARKS: 15', size=12)
    add_paragraph(doc, 'INSTRUCTIONS:', bold=True, size=12)
    add_paragraph(doc, 'ANSWER ALL QUESTIONS', bold=True, size=12)
    doc.add_paragraph()

    q1 = 'a) Assess five characteristics of sociology and explain why they make it a distinct field of study. (5 Marks)'
    a1 = (
        'Sociology is a distinct field of study because it is concerned with the scientific study of society, social relationships and social behaviour. It differs from other disciplines because it studies people in groups, institutions and social systems rather than only as individuals.\n\n'
        'First, sociology is scientific and systematic. It uses observation, surveys, interviews, experiments and critical analysis to gather reliable evidence. This makes it a disciplined field because conclusions are based on facts rather than assumptions or personal opinion.\n\n'
        'Second, sociology studies society as a whole. It focuses on how people live together in families, schools, churches, communities, markets and nations. This broad approach helps explain how social life is organized and how people influence one another.\n\n'
        'Third, sociology studies social institutions. It examines how institutions such as the family, religion, education, government and economy shape behaviour, values and opportunities. These structures have a strong influence on people’s lives and therefore become central to sociological analysis.\n\n'
        'Fourth, sociology is comparative and practical. It compares different societies, groups and historical periods to understand diversity and change. This helps researchers explain why societies differ in culture, inequality, lifestyles and social organization.\n\n'
        'Fifth, sociology is concerned with social problems and social change. It studies poverty, crime, unemployment, migration, gender inequality and political conflict. By analyzing the causes of such problems, sociology provides solutions and informs social policy.\n\n'
        'These characteristics make sociology a distinct field because it uses scientific methods to examine social patterns, institutions and relationships. It is different from other disciplines because it focuses on the collective life of people and how society shapes human behaviour.\n\n'
        'Reference: Giddens, A. (2009). Sociology (6th ed.). Polity Press. Macionis, J. J. (2017). Sociology (15th ed.). Pearson. Kivisto, P. (2013). Social Theory: Roots and Branches (4th ed.). Oxford University Press.'
    )
    add_question(doc, q1, a1)

    q2 = 'b) Discuss the importance of sociology in understanding contemporary Kenyan society. (5 Marks)'
    a2 = (
        'Sociology is very important in understanding contemporary Kenyan society because Kenya is experiencing rapid social, economic and political transformation. It helps explain the behaviour of people and the challenges facing the country in a more meaningful way.\n\n'
        'First, sociology helps explain social problems in Kenya. Issues such as unemployment, poverty, corruption, crime, gender inequality and youth unrest are not individual matters alone; they are shaped by social structures and institutions. Sociology helps us understand their root causes and possible solutions.\n\n'
        'Second, sociology helps explain social inequality. In Kenya, access to resources, education, healthcare, jobs and political opportunities is often influenced by class, gender, ethnicity, religion and region. Sociology enables us to understand how such inequalities are created and sustained.\n\n'
        'Third, sociology helps us understand culture and identity. Kenya is a multi-ethnic and multicultural society with different languages, traditions and social norms. Sociology explains how culture shapes values, social roles and relationships across communities.\n\n'
        'Fourth, sociology is useful in understanding political change and governance. Kenya has experienced reforms, devolution, democratic participation and public debates about accountability. Sociological analysis helps explain how political institutions, citizen participation and state power affect society.\n\n'
        'Fifth, sociology contributes to development planning. Governments, NGOs and development agencies rely on sociological knowledge to design programmes in health, education, agriculture, housing and community development. This ensures policies respond to actual social realities.\n\n'
        'In conclusion, sociology is important in understanding Kenyan society because it explains how people live, interact and organize themselves in a changing world. It provides the knowledge needed to solve social problems and support sustainable national development.\n\n'
        'Reference: Macionis, J. J. (2017). Sociology (15th ed.). Pearson. Giddens, A. (2009). Sociology (6th ed.). Polity Press.'
    )
    add_question(doc, q2, a2)

    q3 = 'c) Discuss three major theories of social change, and explain how they can be applied to Kenya’s ongoing development. (5 Marks)'
    a3 = (
        'Three major theories of social change are evolutionary theory, conflict theory and functionalist theory. These theories explain how societies transform over time and how development occurs. Each theory offers a different interpretation of social change.\n\n'
        'Evolutionary theory argues that societies move gradually from simple to more complex forms. It views social change as progress and modernization. In Kenya, this theory can be applied to the gradual improvement of education, infrastructure, technology and industrial development. It explains how society expands and becomes more advanced over time.\n\n'
        'Conflict theory states that social change occurs because groups with different interests compete for power, resources and opportunities. It argues that change often results from struggle between the powerful and the marginalized. This theory can be applied to Kenya’s development by examining inequalities in land ownership, wealth distribution, political participation and access to services.\n\n'
        'Functionalist theory sees society as an interconnected system in which different institutions work together to maintain stability and order. It suggests that social change occurs when institutions adapt to new needs. In Kenya, this theory can explain reforms in education, healthcare, devolution and governance as efforts to improve social coordination and national development.\n\n'
        'These theories are relevant to Kenya’s ongoing development because the country is undergoing rapid modernization, demographic growth, urbanization and reforms in governance. Evolutionary theory explains modernization, conflict theory explains inequalities and power struggles, while functionalist theory explains institutional adjustment and stability.\n\n'
        'Together, these theories provide a broader understanding of how Kenya is changing and why development is both promising and challenging.\n\n'
        'Reference: Giddens, A. (2009). Sociology (6th ed.). Polity Press. Macionis, J. J. (2017). Sociology (15th ed.). Pearson. Kivisto, P. (2013). Social Theory: Roots and Branches (4th ed.). Oxford University Press.'
    )
    add_question(doc, q3, a3)

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
    doc = build_answer_sheet()
    filename = 'BDS_106_Sociology_of_Development_Answer_Sheet.docx'
    doc.save(filename)
    print(f'Created {filename}')


if __name__ == '__main__':
    main()
