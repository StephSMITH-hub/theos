#!/usr/bin/env python3
"""
generate_dec_study_group_answers.py
Generates comprehensive, publication-grade Study Group Answers for:
STUDY GROUP QUESTIONS (OCTOBER 19TH – 25TH 2026)
CELL LEADERS TRAINING DECEMBER 2021 – FAITHFULNESS IN MINISTRY

STRICTLY sticks to the grammar, word usage, phrasings, illustrations, and logic
of the speaker (Pastor Segun Ayodele) in the teaching, adhering to:
- KeyIndicators KPI Standard and Saints Community Church Study Group Guidelines
- Header Metadata: Name: AGBEDE STEPHEN AYOTOMIWA, Church: SAINTS COMMUNITY CHURCH, AKURE
- Letter page geometry (8.5" x 11.0"), 1.0" margins, Arial 12pt, line spacing 1.15
- Verbatim scripture blockquotes with bold-italic references and verse citations
- Speaker-grounded subheadings originating directly from the teaching's vocabulary
"""

import os
from pathlib import Path
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH

def build_answers(docx_paths, md_paths):
    doc = Document()

    # Page setup: Standard Letter (8.5" x 11.0"), 1.0" margins
    for section in doc.sections:
        section.page_width = Inches(8.5)
        section.page_height = Inches(11.0)
        section.top_margin = Inches(1.0)
        section.bottom_margin = Inches(1.0)
        section.left_margin = Inches(1.0)
        section.right_margin = Inches(1.0)

    # Base style
    style = doc.styles['Normal']
    font = style.font
    font.name = 'Arial'
    font.size = Pt(12)
    font.color.rgb = RGBColor(0x11, 0x18, 0x27) # Dark neutral

    md_lines = []

    def add_title(text):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(2)
        run = p.add_run(text)
        run.bold = True
        run.font.name = 'Arial'
        run.font.size = Pt(18)
        md_lines.append(f"# {text}")

    def add_subtitle(text):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(8)
        run = p.add_run(text)
        run.bold = True
        run.italic = True
        run.font.name = 'Arial'
        run.font.size = Pt(14)
        md_lines.append(f"## *{text}*\n")

    def add_metadata(name, church):
        p1 = doc.add_paragraph()
        p1.paragraph_format.space_before = Pt(0)
        p1.paragraph_format.space_after = Pt(2)
        r1 = p1.add_run(f"Name: {name}")
        r1.bold = True
        r1.font.name = 'Arial'
        r1.font.size = Pt(11)

        p2 = doc.add_paragraph()
        p2.paragraph_format.space_before = Pt(0)
        p2.paragraph_format.space_after = Pt(12)
        r2 = p2.add_run(f"CHURCH: {church}")
        r2.bold = True
        r2.font.name = 'Arial'
        r2.font.size = Pt(11)

        md_lines.append(f"**Name:** {name}  ")
        md_lines.append(f"**CHURCH:** {church}  \n\n---\n")

    def add_q_header(text):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(16)
        p.paragraph_format.space_after = Pt(4)
        run = p.add_run(text)
        run.bold = True
        run.font.name = 'Arial'
        run.font.size = Pt(14)
        md_lines.append(f"## {text}\n")
        return p

    def add_q_prompt(text):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(4)
        p.paragraph_format.space_after = Pt(12)
        run = p.add_run(text)
        run.bold = True
        run.font.name = 'Arial'
        run.font.size = Pt(12)
        md_lines.append(f"**{text}**\n")
        return p

    def add_section_h(text):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(14)
        p.paragraph_format.space_after = Pt(4)
        run = p.add_run(text)
        run.bold = True
        run.font.name = 'Arial'
        run.font.size = Pt(13)
        md_lines.append(f"### {text}\n")
        return p

    def add_body_subh(text):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(12)
        p.paragraph_format.space_after = Pt(4)
        run = p.add_run(text)
        run.bold = True
        run.font.name = 'Arial'
        run.font.size = Pt(12)
        md_lines.append(f"#### {text}\n")
        return p

    def ref(text):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(8)
        p.paragraph_format.space_after = Pt(2)
        run = p.add_run(text)
        run.bold = True
        run.italic = True
        run.font.name = 'Arial'
        run.font.size = Pt(12)
        md_lines.append(f"***{text}***")
        return p

    def v(text):
        p = doc.add_paragraph()
        p.paragraph_format.left_indent = Inches(0.25)
        p.paragraph_format.space_before = Pt(1)
        p.paragraph_format.space_after = Pt(1)
        run = p.add_run(text)
        run.bold = True
        run.italic = True
        run.font.name = 'Arial'
        run.font.size = Pt(12)
        md_lines.append(f"> ***{text}***  ")
        return p

    def sp():
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(3)
        p.paragraph_format.space_after = Pt(3)
        md_lines.append("")
        return p

    def b(runs_data):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(4)
        p.paragraph_format.space_after = Pt(4)
        p.paragraph_format.line_spacing = 1.15
        md_para = ""
        for r_text, r_bold, r_italic in runs_data:
            run = p.add_run(r_text)
            run.font.name = 'Arial'
            run.font.size = Pt(12)
            if r_bold:
                run.bold = True
            if r_italic:
                run.italic = True

            seg = r_text
            if r_bold and r_italic:
                seg = f"***{seg}***"
            elif r_bold:
                seg = f"**{seg}**"
            elif r_italic:
                seg = f"*{seg}*"
            md_para += seg

        md_lines.append(md_para + "\n")
        return p

    # =========================================================================
    # DOCUMENT HEADER
    # =========================================================================
    add_title("STUDY GROUP QUESTIONS (OCTOBER 19TH – 25TH 2026)")
    add_subtitle("CELL LEADERS TRAINING DECEMBER 2021 – FAITHFULNESS IN MINISTRY")
    add_metadata("AGBEDE STEPHEN AYOTOMIWA", "SAINTS COMMUNITY CHURCH, AKURE")

    # =========================================================================
    # QUESTION 1
    # =========================================================================
    add_q_header("QUESTION 1")
    add_q_prompt("Explain the difference between faithfulness and ambition using the examples of Ruth and Naomi, Elisha and Elijah, Joshua and Moses, Timoth and Paul as done within this track.")

    add_section_h("INTRODUCTION")

    b([
        ("In the local church and the systematic discipling of Christian workers, the governing requirement laid upon every minister is singularly defined in ", False, False),
        ("1 Corinthians 4:1-2", True, False),
        (": ", False, False),
        ("'Let a man so account of us, as of the ministers of Christ, and stewards of the mysteries of God. Moreover it is required in stewards, that a man be found faithful.' ", True, True),
        ("In this teaching, Pastor establishes from scratch the exact definition of what it means to be faithful. The word faithful is the same word for faith, ", False, False),
        ("pisteuo", True, True),
        (". It means to be reliable, to be trustworthy, and to be dependable under every circumstance. When Scripture declares in ", False, False),
        ("1 Corinthians 10:13", True, False),
        (" that ", False, False),
        ("God is faithful", True, True),
        (", it is because He is reliable, He is always available, and He does not change His commitment or alter His word in spite and despite opposing circumstances or extremities. Therefore, as taught in this track, ", False, False),
        ("faithfulness is also called loyalty", True, False),
        (", and ", False, False),
        ("loyalty means to be committed to a cause 100%", True, True),
        (".", False, False)
    ])

    b([
        ("From this foundation, Pastor draws a sharp, irreconcilable distinction between ", False, False),
        ("faithfulness", True, False),
        (" and ", False, False),
        ("ambition", True, False),
        (". There are people who join a church or enter ministry specifically so that they can rise through the ranks, gain visibility, or be noticed by leadership. The teaching provides the definitive spiritual diagnostic to unmask mere ambition: ", False, False),
        ("one of the ways you know whether you are faithful or you are just ambitious is if your desires are not given to you", True, True),
        (". The ambitious person will act faithful only as long as the trajectory favors their personal desires; any little alteration, delay, or denial of their expectations, and they will change immediately. That sudden change proves they were never faithful. Faithful men, however, do not break under pressure; faithfulness has nothing to do with circumstances, because the faithful steward makes no excuses and remains 100% committed to the cause. To thoroughly demonstrate this difference, the teaching examines four biblical relationships: ", False, False),
        ("Ruth and Naomi", True, False),
        (" (uncalculating loyalty that looks for nothing), ", False, False),
        ("Elisha and Elijah", True, False),
        (" (wholehearted servanthood and wearing the father's ministry), ", False, False),
        ("Joshua and Moses", True, False),
        (" (staying with the leader's instructions without preaching the leader's flaws), and ", False, False),
        ("Timothy and Paul", True, False),
        (" (faithful dependability everywhere, contrasted with the unfaithful man who is like a broken tooth and a foot out of joint).", False, False)
    ])

    sp()
    add_section_h("BODY")

    add_body_subh("The Meaning of Faithfulness and the Litmus Test of Ambition")

    ref("1 Corinthians 4:1-2")
    v("[1] Let a man so account of us, as of the ministers of Christ, and stewards of the mysteries of God.")
    v("[2] Moreover it is required in stewards, that a man be found faithful.")
    sp()

    ref("1 Corinthians 10:13")
    v("[13] There hath no temptation taken you but such as is common to man: but God is faithful, who will not suffer you to be tempted above that ye are able...")
    sp()

    b([
        ("As expounded in the teaching, the Apostle Paul declares that a Christian worker is a steward, and the primary qualification demanded of a steward is that they be found ", False, False),
        ("faithful", True, False),
        (". The speaker explains that the word faithful is rooted in ", False, False),
        ("pisteuo", True, True),
        (", signifying that which is utterly reliable and trustworthy. Believers often say, 'God is faithful.' Why is God faithful? Because He is reliable, He is always available, and He never changes His commitment. Look at what God walked through before Jesus died on the cross—it was a lot of negativities, extremities, and intense pressure, yet God never altered His word. That is why the speaker emphasizes that ", False, False),
        ("faithfulness is loyalty, and loyalty means being committed to a cause 100%", True, False),
        (". A faithful steward is not loyal because conditions are favorable; they are reliable under every circumstance.", False, False)
    ])

    b([
        ("In sharp contrast stands ", False, False),
        ("personal ambition", True, False),
        (". The speaker reveals how ambition operates within the church: individuals attach themselves to a work or join a fellowship with a hidden agenda—to climb ecclesiastical ladders, gain a platform, or be placed in visible authority. Because their motive is self-advancement rather than the cause, they are merely ", False, False),
        ("'acting faithful.'", True, True),
        (" The speaker gives the exact way to know their true condition: ", False, False),
        ("'One of the ways you know whether you are faithful or you're just ambitious is if your desires are not given you, they're not given to you... Any little alteration, they will change. That means they were never faithful.'", True, True),
        (" When an ambitious worker is corrected, when their preferred promotion is delayed, or when another worker is chosen, their disposition changes overnight. They begin to murmur, withdraw their service, or leave the church entirely. Faithful people, however, do not break under pressure; they possess no personal entitlement, and their loyalty remains constant whether their natural desires are fulfilled or denied.", False, False)
    ])

    add_body_subh("Ruth and Naomi: Uncalculating Loyalty Looking for Nothing")

    ref("Ruth 1:14, 16-17")
    v("[14] And they lifted up their voice, and wept again: and Orpah kissed her mother in law; but Ruth clave unto her.")
    v("[16] And Ruth said, Intreat me not to leave thee, or to return from following after thee: for whither thou goest, I will go; and where thou lodgest, I will lodge: thy people shall be my people, and thy God my God:")
    v("[17] Where thou diest, will I die, and there will I be buried: the LORD do so to me, and more also, if ought but death part thee and me.")
    sp()

    b([
        ("The first illustration employed in the teaching to demonstrate the difference between faithfulness and ambition is ", False, False),
        ("Ruth and Naomi", True, False),
        (". Following the death of her husband and both sons in the country of Moab, Naomi stood utterly bereft of social security, financial inheritance, and prospective husbands. Recognizing her destitution, Naomi urged her two Moabite daughters-in-law to return to their mothers' houses. ", False, False),
        ("Orpah kissed Naomi and departed, but Ruth clave unto her", True, False),
        (" (Ruth 1:14). Orpah exemplifies affection that terminates when personal advantage ceases; she loved Naomi, but once Naomi had no more sons or worldly prospects to offer, Orpah turned back.", False, False)
    ])

    b([
        ("Ruth, however, embodies pure, unadulterated faithfulness. When urged to leave, Ruth uttered that monumental covenant of devotion: ", False, False),
        ("'Whither thou goest, I will go; and where thou lodgest, I will lodge: thy people shall be my people, and thy God my God: Where thou diest, will I die, and there will I be buried' ", True, True),
        ("(Ruth 1:16-17). The speaker captures the entire essence of Ruth's action in one definitive sentence: ", False, False),
        ("'She wasn't looking for anything. That's faithfulness.'", True, True),
        (" Ruth was an alien Gentile entering Israel with zero guarantees of marriage, property, or societal respect. She was not calculating perks, perks of leadership, or future benefits. Ambition is always calculating: 'What will I get out of this association?' Faithfulness declares: 'I am committed 100% to this spiritual cause and this oversight, looking for nothing in return.'", False, False)
    ])

    add_body_subh("Elisha and Elijah: Wholehearted Servanthood and Wearing the Father's Ministry")

    ref("2 Kings 3:11")
    v("[11] But Jehoshaphat said, Is there not here a prophet of the LORD, that we may enquire of the LORD by him? And one of the king of Israel's servants answered and said, Here is Elisha the son of Shaphat, which poured water on the hands of Elijah.")
    sp()

    ref("2 Kings 2:2, 14")
    v("[2] And Elijah said unto Elisha, Tarry here, I pray thee; for the LORD hath sent me to Bethel. And Elisha said unto him, As the LORD liveth, and as thy soul liveth, I will not leave thee. So they went down to Bethel.")
    v("[14] And he took the mantle of Elijah that fell from him, and smote the waters, and said, Where is the LORD God of Elijah? and when he also had smitten the waters, they parted hither and thither: and Elisha went over.")
    sp()

    b([
        ("The second example analyzed in the teaching is the relationship between ", False, False),
        ("Elisha and Elijah", True, False),
        (". The speaker highlights the historical background: Elisha emerged from a wealthy family (plowing with twelve yoke of oxen) and was likely older than Elijah, yet he completely surrendered his independence to follow Elijah faithfully. Years later, when the kings of Israel, Judah, and Edom were in crisis and inquired for a prophet of the Lord, Elisha was identified not by ecclesiastical titles or prophetic claims, but by his steadfast servanthood: ", False, False),
        ("'Here is Elisha the son of Shaphat, which poured water on the hands of Elijah' ", True, True),
        ("(2 Kings 3:11). The speaker notes that pouring water on Elijah's hands was a symbolic term meaning ", False, False),
        ("he served him faithfully with all his heart", True, False),
        (".", False, False)
    ])

    b([
        ("The speaker contrasts Elisha's faithful devotion with the behavior of the ", False, False),
        ("'sons of the prophets'", True, False),
        (". As Elijah and Elisha pressed through Gilgal, Bethel, Jericho, and toward Jordan, the sons of the prophets stayed behind at a distance, taunting Elisha with cynical questions: ", False, False),
        ("'Don't you know? Why are you following this man? What are you following him for? Don't you know he's going today?'", True, True),
        (" The sons of the prophets represent onlookers who possess information about the prophetic office but lack relational loyalty and the willingness to endure pressure. Elisha did not even answer their cynicism; he stood faithful, declaring repeatedly, ", False, False),
        ("'As the LORD liveth, and as thy soul liveth, I will not leave thee.'", True, True)
    ])

    b([
        ("When Elijah asked what he should do for him, Elisha asked for a double portion of his spirit. The speaker clarifies that this was ", False, False),
        ("the firstborn inheritance to be the head of the prophets", True, False),
        ("—taking on the spiritual responsibility of continuing Elijah's work, not an ambitious attempt to outshine him. The crowning proof that Elisha was not acting out of personal ambition occurred the moment Elijah was caught up into heaven. Picking up the fallen mantle, Elisha struck the waters of Jordan and cried: ", False, False),
        ("'Where is the LORD God of Elijah?'", True, True),
        (" The speaker powerfully underscores this exact phrase: ", False, False),
        ("'That's what he said. That's faithfulness. Nothing changed. He wasn't doing it for himself. Where is the Lord God? He didn't say of Elisha? No, he said of Elijah. So he's letting everybody know: I'm wearing Elijah's ministry. That's a faithful person. He's loyal.'", True, True),
        (" Ambition seeks to establish an independent brand: 'Now that the old man is gone, my time has come; behold the ministry of Elisha!' Faithfulness honors spiritual lineage and lets everyone know: 'I am merely wearing my father's ministry under the same God.'", False, False)
    ])

    add_body_subh("Joshua and Moses: Staying with Instructions and Never Preaching with the Leader's Mistakes")

    ref("Joshua 1:1-2")
    v("[1] Now after the death of Moses the servant of the LORD it came to pass, that the LORD spake unto Joshua the son of Nun, Moses' minister, saying,")
    v("[2] Moses my servant is dead; now therefore arise, go over this Jordan, thou, and all this people, unto the land which I do give to them, even to the children of Israel.")
    sp()

    b([
        ("The third case study is ", False, False),
        ("Joshua and Moses", True, False),
        (". Throughout the wilderness journey, Joshua was identified as ", False, False),
        ("Moses' minister", True, False),
        (" (Exodus 24:13; Joshua 1:1)—a dependable assistant who executed Moses' commands and remained in the tabernacle. When Moses died and God called Joshua to lead the children of Israel across Jordan, Joshua's loyalty was put to the ultimate test. How does an ambitious person behave when succeeding a legendary predecessor? The speaker exposes the subtle carnality of ambition: an ambitious successor frequently uses the predecessor's flaws to build their own credibility and preach to the congregation.", False, False)
    ])

    b([
        ("The speaker explains the contrast in vivid, unmistakable language: ", False, False),
        ("'Joshua said, this is the God of Moses as Moses, the servant of the Lord. And you know, he didn't use Moses to preach, to say, you know, Moses could have brought us here, but you know he's not the one that eventually came, you know what happened now? Should you get? No, no, no, no. He stayed with the fact that this is Moses' instruction. Loyal, faithful people don't do otherwise when you are not there.'", True, True),
        (" An ambitious leader exploits the leader's absence or passing to highlight where the predecessor failed, presenting themselves as the superior, enlightened leader. Joshua completely refused to preach with Moses' mistakes. He stayed strictly with Moses' instructions, honored Moses as the servant of the Lord, and carried the people across Jordan. Faithfulness preserves instructions; ambition alters them for personal glory.", False, False)
    ])

    add_body_subh("The Behind-the-Scenes Standard: The Front-Row Test")

    b([
        ("Building directly upon Joshua's loyalty in Moses' absence, Pastor delivers the foundational litmus test that separates genuine faithfulness from mere eye-service in the local church: ", False, False),
        ("'Loyal, faithful people don't do otherwise when you are not there. Let me add one more. Faithful people will do in the service what they will have done if Jesus was seated in the front seat. Faithful people will do in the service what they will have done if Pastor was seated in the service. Faithful people, it is not because you are there. In fact, they do things behind you as though you are seated there. It's required of a steward that a man be found faithful, reliable under every circumstance.'", True, True)
    ])

    b([
        ("Ambitious individuals are governed entirely by visibility: they arrive early, pray loudly, take detailed notes, and display fervent activity only when the Senior Pastor or coordinator is physically present in the room. The moment leadership is absent, their standard collapses into flippancy, lateness, and negligence. The faithful steward, however, operates under the perpetual consciousness of God. Whether laboring in front of thousands or organizing chairs in complete obscurity, their conduct is identical to how they would behave if Jesus Christ or their pastor were seated right in the front row.", False, False)
    ])

    add_body_subh("Timothy and Paul: Faithful Everywhere vs. The Broken Tooth and Foot Out of Joint")

    ref("Proverbs 25:19")
    v("[19] Confidence in an unfaithful man in time of trouble is like a broken tooth, and a foot out of joint.")
    sp()

    ref("Philippians 2:19-22")
    v("[19] But I trust in the Lord Jesus to send Timotheus shortly unto you, that I also may be of good comfort, when I know your state.")
    v("[20] For I have no man likeminded, who will naturally care for your state.")
    v("[21] For all seek their own, not the things which are Jesus Christ's.")
    v("[22] But ye know the proof of him, that, as a son with the father, he hath served with me in the gospel.")
    sp()

    b([
        ("The fourth example presented in the teaching is ", False, False),
        ("Timothy and Paul", True, False),
        (". Drawing from what Pastor Somi shared at the Believers Convention, the speaker notes how thoroughly dependable Timothy was across Paul's entire apostolic journey: ", False, False),
        ("'And we saw Timothy everywhere, you know, we can readily call Paul's epistles, Timothy's epistles, readily faithful. Faithful. He's everywhere, he's everywhere.'", True, True),
        (" In almost every epistle written by Paul—from Thessalonians to Corinthians, Philippians, and Colossians—Timothy is present as a trusted co-laborer. In Philippians 2:20-22, Paul lamented that while most ministers seek their own interests, Timothy served with him as a son with a father and naturally cared for the state of the flock.", False, False)
    ])

    b([
        ("The speaker then presents the exact scriptural contrast to Timothy's faithfulness by reading ", False, False),
        ("Proverbs 25:19", True, False),
        (": ", False, False),
        ("'Confidence in an unfaithful man in time of trouble is like a broken tooth, and a foot out of joint.'", True, True),
        (" The speaker expounds this verse with visceral practical application: ", False, False),
        ("'I've experienced it too. A broken tooth: it will be painful, it cannot chew. A foot out of joint: you can't walk with it, it will not go. He's there. He has a responsibility. You have made him a cell leader, he's there. His name is on the roll, he's HOD, he's a pastor. But it's like a broken tooth: we can't chew, we can't walk, but he's there.'", True, True)
    ])

    b([
        ("An ambitious, unfaithful worker is like a dislocated ankle or a fractured molar: in calm weather, they appear functional, but the moment trouble, pressure, or demands arrive, leaning on them produces excruciating pain and ministerial paralysis. They occupy an office, draw a title, and appear on the roster, yet when work is to be done, they cannot chew and cannot walk. Timothy was the exact opposite: steady, reliable, and available under every condition. The speaker delivers a sobering warning: ", False, False),
        ("'Anybody that tells you that what you do in the service of God has no consequence, he's deceiving you... When you are serving God, you don't joke. God practically treasures lives.'", True, True)
    ])

    sp()
    add_section_h("CONCLUSION")

    b([
        ("In conclusion, the difference between faithfulness and ambition within this teaching is total and absolute. ", False, False),
        ("Ambition", True, False),
        (" is driven by self-advancement, title-seeking, and rising through the ranks. It acts faithful only as long as its desires are granted; the moment an alteration or denial occurs, it changes, murmurs, and withdraws. When trouble arrives, the ambitious worker is like a broken tooth that cannot chew and a foot out of joint that cannot walk—paralyzing the local church.", False, False)
    ])

    b([
        ("Faithfulness", True, False),
        (", conversely, is ", False, False),
        ("loyalty—being committed to a cause 100%", True, True),
        (". It is reliable, trustworthy, and constant regardless of circumstances. Faithful men do not break under pressure, make no excuses, and remain immovable. It is embodied in: ", False, False),
        ("Ruth", True, False),
        (", who clave to Naomi looking for nothing; ", False, False),
        ("Elisha", True, False),
        (", who poured water on Elijah's hands, ignored the cynical sons of the prophets, and struck Jordan crying, 'Where is the Lord God of Elijah?', letting everyone know he was wearing his father's ministry; ", False, False),
        ("Joshua", True, False),
        (", who stayed with Moses' instructions and refused to preach with Moses' flaws; and ", False, False),
        ("Timothy", True, False),
        (", who was everywhere faithful, serving as a son with a father.", False, False)
    ])

    b([
        ("For every cell leader and worker, the challenge is clear: We must examine our motives before God. We must pass the front-row test—doing in the service and behind the scenes precisely what we would do if Jesus Christ or our Pastor were seated in the front seat. We must discard personal ambition, crucify self-seeking entitlements, and be found faithful stewards who are 100% dependable in the house of God.", False, False)
    ])

    # =========================================================================
    # QUESTION 2
    # =========================================================================
    add_q_header("QUESTION 2")
    add_q_prompt("Does God test believers? With the aid of this track explain the correlation between tests and faithfulness in the work of the ministry. Also explain the role of tests in attaining leadership. Using at least 2 biblical characters employed in this teaching, explain the possible outcomes to tests.")

    add_section_h("INTRODUCTION")

    b([
        ("A critical doctrinal question confronted in this teaching is whether God tests believers under the New Covenant. The speaker answers with an emphatic, unqualified ", False, False),
        ("YES", True, False),
        (". Pastor states directly: ", False, False),
        ("'Praise the Lord. And there are things that will come to test you. To test that you are faithful, to test you. And some of us have failed the test. You never know when the test is. God tests people.'", True, True),
        (" However, the teaching immediately clarifies the exact nature of divine testing: ", False, False),
        ("'It is not testing you so that you can fail. It's the way you test drive a car.'", True, True),
        (" Just as an automobile manufacturer test-drives a newly engineered vehicle through demanding terrains and braking tests not to wreck the engine, but to prove that it can handle real-world road conditions without breaking down, God tests believers to authenticate their integrity and build their capacity.", False, False)
    ])

    b([
        ("The speaker points to God's boast over Job in ", False, False),
        ("Job 1:8", True, False),
        (": ", False, False),
        ("'Have you considered my servant Job? What a boast! Have you considered him? He has integrity, that test.'", True, True),
        (" In the same vein, ", False, False),
        ("James 1:2-4", True, False),
        (" commands believers to count it all joy when they encounter diverse tests and trials, because the trial of faith produces patience (", False, False),
        ("hypomone", True, True),
        ("), and that patience, when finished, leaves the minister entirely wanting nothing. Pastor delivers a stern warning: ", False, False),
        ("'When you run from tests, you run from the character that you require to do the work of the ministry.'", True, True),
        (" Tests are not an interruption of ministry; they are the divine instrument through which faithfulness is forged, character is proven, and leaders are qualified.", False, False)
    ])

    b([
        ("To provide a comprehensive exposition of Question 2 as done within this track, this answer analyzes: (1) the divine purpose of tests as illustrated by the car test-drive, (2) the direct correlation between tests and faithfulness in ministry, (3) the mandatory role of human proving before appointing leaders in the local church (1 Tim 3:10, 13; 2 Tim 2:15; Acts 15:26), and (4) the divergent outcomes of testing through the primary archetypes of ", False, False),
        ("Joseph", True, False),
        (" (passing into Christlike character and kingdom authority) versus ", False, False),
        ("Gehazi", True, False),
        (" (failing through covetousness into leprosy and ruin), supplemented by ", False, False),
        ("Eliab", True, False),
        (" (who failed the test of the heart) and ", False, False),
        ("David", True, False),
        (" (who passed the test of the few sheep).", False, False)
    ])

    sp()
    add_section_h("BODY")

    add_body_subh("Does God Test Believers? The Purpose of the Divine Test-Drive")

    ref("James 1:2-4")
    v("[2] My brethren, count it all joy when ye fall into divers temptations;")
    v("[3] Knowing this, that the trying of your faith worketh patience.")
    v("[4] But let patience have her perfect work, that ye may be perfect and entire, wanting nothing.")
    sp()

    ref("Job 1:8")
    v("[8] And the LORD said unto Satan, Hast thou considered my servant Job, that there is none like him in the earth, a perfect and an upright man, one that feareth God, and escheweth evil?")
    sp()

    b([
        ("The teaching leaves no ambiguity regarding whether God tests believers: ", False, False),
        ("'God tests people. It is not testing you so that you can fail. It's the way you test drive a car.'", True, True),
        (" When an automobile company designs a vehicle, the car is subjected to rigorous crash tests, high-speed tracks, and severe braking. The engineers do not test-drive the vehicle to crush it or cause it to fail; they test-drive it to prove its roadworthiness and demonstrate that the mechanical components will hold up under severe stress. In the same way, God tests believers to prove their integrity, burn out instability, and display their dependability.", False, False)
    ])

    b([
        ("The speaker notes how God boasted to Satan concerning Job: ", False, False),
        ("'Have you considered my servant Job? What a boast. Have you considered him? He has integrity, that test.'", True, True),
        (" God knew Job's faith was not fair-weather enthusiasm; his integrity was tested and proven. James 1:2-4 confirms that the trial of faith produces patience—the capacity to stand firm under pressure. Therefore, ", False, False),
        ("'when you run from tests, you run from the character that you require to do the work of the ministry.'", True, True),
        (" A Christian worker who evades difficulty, runs from correction, or abandons their post whenever friction arises will never develop the spiritual stamina required to shepherd people.", False, False)
    ])

    add_body_subh("The Correlation Between Tests and Faithfulness in the Work of Ministry")

    b([
        ("Within the teaching, testing and faithfulness are shown to be completely inseparable. You cannot claim to be faithful in the abstract; faithfulness can only be verified when tested under pressure. As Pastor states: ", False, False),
        ("'Faithful men don't break under pressure. No, he's faithful, reliable. No excuses.'", True, True),
        (" Tests are the divine mechanism that reveals whether a leader is truly committed to the cause or merely enjoying the comfort of convenience. God tests you to prove that you are faithful, and ", False, False),
        ("'God does not use men that He doesn't test. He will use you to test you so that He can use you for more. That's what He does.'", True, True)
    ])

    b([
        ("The speaker shares a deeply personal reflection on how a mature leader interprets trials: ", False, False),
        ("'As I am right now, sometimes I look at myself and say, this is God working something through me. He's building character. He's building strength, commitment. That's what He's doing. So you can read meanings to different circumstances. See something as a trial for you. Or see it like, I'm tired, I'm tired. No, you shouldn't be. You have to be tested.'", True, True),
        (" When a cell leader views challenges through the eyes of flesh, they complain about fatigue, lack of transport, or bad weather. But when viewed through spiritual maturity, every trial is recognized as God's gym—building ministerial muscles, endurance, and unyielding dependability.", False, False)
    ])

    b([
        ("Furthermore, the teaching establishes that ", False, False),
        ("'faithfulness is the breath of ministry that you are always at your duty post.'", True, True),
        (" The speaker uses the practical example of ", False, False),
        ("holidays", True, False),
        (" to demonstrate ministerial alertness. The speaker declares: ", False, False),
        ("'The holiday period is not holiday for ministry. No, it's not. That is why I grew up saying I don't take holidays from the government... On Christmas Day, I made observance because I like to observe: in plane, security people were working, the planes were working, air traffic controller didn't say jingle bells... Doctors are working, they will call on Christmas Day... Things are working around the clock. The police are saying Merry Christmas, please let me check your boot. Everybody is working! So I told myself, I'm a preacher of the gospel, I must be alert. So you must be faithful with time, faithful with alertness in the things that you are meant to do.'", True, True),
        (" Tests reveal whether a leader maintains unbroken vigilance at their post or abandons the flock the moment a secular holiday or personal fatigue arrives.", False, False)
    ])

    add_body_subh("The Role of Tests in Attaining Leadership: The Mandate of Proving")

    ref("1 Timothy 3:10, 13")
    v("[10] And let these also first be proved; then let them use the office of a deacon, being found blameless.")
    v("[13] For they that have used the office of a deacon well purchase to themselves a good degree, and great boldness in the faith which is in Christ Jesus.")
    sp()

    ref("2 Timothy 2:15")
    v("[15] Study to shew thyself approved unto God, a workman that needeth not to be ashamed, rightly dividing the word of truth.")
    sp()

    ref("Acts 15:25-26")
    v("[25] It seemed good unto us, being assembled with one accord, to send chosen men unto you with our beloved Barnabas and Paul,")
    v("[26] Men that have hazarded their lives for the name of our Lord Jesus Christ.")
    sp()

    b([
        ("In the local church, testing is the non-negotiable prerequisite for appointing leaders. The speaker lays down a cardinal rule for all pastors and coordinators: ", False, False),
        ("'Don't use people you have not proven. Don't dash, don't donate responsibility to people.'", True, True),
        (" Reading from 1 Timothy 3:10, the speaker points out: ", False, False),
        ("'Let these also first be proved. The Bible says let them be examined closely, tested thoroughly.'", True, True),
        (" The speaker then asks a penetrating question that shatters religious passivity: ", False, False),
        ("'If Paul writes to Timothy that he should prove the deacons, who does the proving? Huh? Is he saying the Spirit of God will prove him? No, you will do it! So in leadership you open your eyes to moments to know people and their true character. It must be proved.'", True, True)
    ])

    b([
        ("Spiritual leaders cannot wait for a mystical voice or an angelic certificate from heaven; human leadership must actively prove prospective workers through real-life observations. In 2 Timothy 2:15, Paul says to be ", False, False),
        ("'approved unto God'—approved is the exact same word: tested! A leader cannot be untested.", True, True),
        (" In Acts 15:26, when the apostles in Jerusalem sent Paul and Barnabas to the Gentile churches, they commended them as: ", False, False),
        ("'Men that have hazarded their lives for the name of our Lord Jesus Christ. You know what that means? They put everything on the table. Said, we know these guys, we've seen them through hard times. They are not fair-weather guys. So he said, listen to them. They have been proven.'", True, True)
    ])

    b([
        ("The speaker details exactly how testing is practically conducted in the local church to vet leadership and weed out ambition:", False, False)
    ])

    b([
        ("1. Pocketing Titles at the Door: ", True, False),
        ("When people join the church seeking rapid prominence, their titles mean nothing: ", False, False),
        ("'Wherever you are coming from, even if you are an Archbishop, you pocket it when you get to the door here! You join membership class, we see you every day, we test you with Operations department, then you go to Works, then Security... Those who are singing, they now wash toilet! You come here to have position or to serve? If you are here to serve, anywhere they put you, you will serve.'", True, True),
        (" Ambitious people flee when handed a mop or assigned to security; true servants serve with joy wherever placed.", False, False)
    ])

    b([
        ("2. The Zonal Pastor Reporting Test: ", True, False),
        ("The speaker recounts an ambitious worker who claimed he was gathering people in church: ", False, False),
        ("'Someone told me, say, Pastor, I'm gathering some people... I said, yeah, go and do it in your zone, submit the report to the Zonal Pastor. Person became uncomfortable. I knew he was ambitious. Person doesn't come here anymore, because that was the motive—so that Pastor will notice me!'", True, True),
        (" Testing motives through established church structures immediately exposes self-promoters.", False, False)
    ])

    b([
        ("3. Practical Operational Testing: ", True, False),
        ("Leadership observes workers in real responsibilities: ", False, False),
        ("'We put you in the committee for Camp Meeting or Believers Convention, and the hostel you're supposed to supervise went bad, we are going to be looking at that. Why? You're supposed to handle the Works department, you got there late—they matter! They do matter. You had eight workers working with you, two of them remain, the other six we don't see them anymore—they matter a lot. We look back and say, you don't look like you're ready for this. Eye-service can't take you far.'", True, True)
    ])

    b([
        ("4. Fires in the Desert: ", True, False),
        ("Pastor explains why trials are necessary before leadership promotion: ", False, False),
        ("'Sometimes God will allow some fires so that those who cannot be with you in the Promised Land should not cross the desert. They shouldn't. So there are fires that you must have... God doesn't give oversight by social media. Forget it, it is a lie.'", True, True),
        (" Testing ensures that unfaithful, brittle individuals are not entrusted with the flock of God.", False, False)
    ])

    add_body_subh("Biblical Characters and Possible Outcomes to Tests: Joseph vs. Gehazi")

    b([
        ("The teaching explicitly summarizes the two divergent outcomes of testing in one clear statement: ", False, False),
        ("'So when you test people, they come out either as Joseph or as Gehazi.'", True, True),
        (" Passing tests produces Joseph-like Christlike character and kingdom exaltation; failing tests produces Gehazi-like compromise, disqualification, and spiritual leprosy.", False, False)
    ])

    ref("Genesis 39:9; Genesis 50:20")
    v("[9] There is none greater in this house than I; neither hath he kept back any thing from me but thee, because thou art his wife: how then can I do this great wickedness, and sin against God?")
    v("[20] But as for you, ye thought evil against me; but God meant it unto good, to bring to pass, as it is this day, to save much people alive.")
    sp()

    b([
        ("Outcome 1: Passing the Test — Joseph (Endurance, Christlike Character, and Exaltation):", True, False)
    ])

    b([
        ("Joseph stands as the supreme example of a man who passed every escalating test across every station of life. The speaker traces his faithfulness through four major environments:", False, False)
    ])

    b([
        ("• In Potiphar's House as a Slave: ", True, False),
        ("Joseph went from being his father's favored son into Egyptian slavery. Yet, ", False, False),
        ("'when you are faithful, no circumstance changes you. He got to Potiphar's house, he's faithful. Potiphar committed the whole place to his hands.'", True, True),
        (" When Potiphar's wife repeatedly solicited him to commit adultery, Joseph refused to compromise, asserting: ", False, False),
        ("'How then can I do this great wickedness, and sin against God?' (Genesis 39:9).", True, True),
        (" The speaker calls this ", False, False),
        ("'the high point of that faithfulness.'", True, False),
        (" She tore his clothes and defamed him, yet Joseph stayed faithful.", False, False)
    ])

    b([
        ("• In Pharaoh's Prison: ", True, False),
        ("Falsely accused and imprisoned, Joseph did not collapse into self-pity, bitterness, or rebellion: ", False, False),
        ("'He didn't get to prison and he's just grumbling: After serving God, everything I tried to do... let me cuckoo finish it, what's the big deal, let's plan prison escape. No, he didn't!'", True, True),
        (" Nor did he grumble when other prisoners had dreams: ", False, False),
        ("'He wasn't bitter and said, is it not dream that got me here? Look, I'm not ready for all this ministry nonsense... I just want to get out and live my life and move on. No, he just interpreted it. Faithful, faithful, faithful!'", True, True),
        (" Even when the butler forgot him for two full years, Joseph nurtured no bitterness.", False, False)
    ])

    b([
        ("• Before Pharaoh: ", True, False),
        ("The speaker notes: ", False, False),
        ("'If you are going to handle a nation, you must know how to handle trouble. So he had learned how to handle difficulties.'", True, True),
        (" When summoned before Pharaoh, Joseph did not bargain or set self-centered conditions: ", False, False),
        ("'He didn't get there and say, I can't interpret your dream unless conditions: one, they will release me; two, that woman that lied against me, I want her head; three, I want to go back home. He just interpreted the dream, and the man said, it has to be you, the Spirit of God is in you.'", True, True)
    ])

    b([
        ("• In Ultimate Authority: ", True, False),
        ("Exalted over Egypt, Joseph governed the entire nation—including Potiphar's household: ", False, False),
        ("'When she needs food, she will line up say, Auntie Potty! Say, sorry my lord. Don't worry, it is well, give her Indomie, macaroni and all that, let her go. He has learned it!'", True, True),
        (" When his fearful brothers came after Jacob died, concocting a text message out of fear, Joseph comforted them, declaring: ", False, False),
        ("'What you meant for evil, God meant for good... See the character of Jesus Christ! That's faithfulness. No circumstance changes a faithful man.'", True, True),
        (" Passing the test forged in Joseph the unmistakable character of Christ.", False, False)
    ])

    sp()
    ref("2 Kings 5:20, 26-27")
    v("[20] But Gehazi, the servant of Elisha the man of God, said, Behold, my master hath spared Naaman this Syrian, in not receiving at his hands that which he brought: but, as the LORD liveth, I will run after him, and take somewhat of him.")
    v("[26] And he said unto him, Went not mine heart with thee, when the man turned again from his chariot to meet thee? Is it a time to receive money, and to receive garments...?")
    v("[27] The leprosy therefore of Naaman shall cleave unto thee, and unto thy seed for ever. And he went out from his presence a leper as white as snow.")
    sp()

    b([
        ("Outcome 2: Failing the Test — Gehazi (Covetousness, Compromise, and Leprosy):", True, False)
    ])

    b([
        ("In tragic contrast to Joseph, the speaker presents ", False, False),
        ("Gehazi", True, False),
        (" as the archetype of the minister who fails the test: ", False, False),
        ("'As a minister, you could be like Gehazi. Gehazi wanted something. He went after Naaman and he got leprosy. He became leprous. Test!'", True, True),
        (" Gehazi was the personal associate of the prophet Elisha. He had witnessed supernatural power, dead children raised, and miracles multiplied. He was in prime position to inherit Elisha's mantle, just as Elisha had inherited from Elijah. But Gehazi's heart was untested regarding material wealth.", False, False)
    ])

    b([
        ("When Naaman the Syrian general was healed of leprosy and offered lavish gifts, Elisha adamantly refused remuneration, preserving the pure grace of God. But Gehazi's hidden ambition and covetousness flared. Swearing a carnal oath, Gehazi ran after Naaman, fabricated a lie about visiting sons of the prophets, collected two talents of silver and changes of garments, and hid them in his house. Elisha confronted him by the Spirit: ", False, False),
        ("'Went not mine heart with thee, when the man turned again from his chariot to meet thee? Is it a time to receive money, and to receive garments... The leprosy therefore of Naaman shall cleave unto thee, and unto thy seed for ever.'", True, True),
        (" Gehazi walked out a leper as white as snow. Proximity to an anointed ministry did not save him. His untested, greedy heart failed the test, resulting in permanent disqualification, ruin, and spiritual leprosy.", False, False)
    ])

    add_body_subh("Supplementary Biblical Characters: Eliab vs. David")

    ref("1 Samuel 16:6-7, 11")
    v("[6] And it came to pass, when they were come, that he looked on Eliab, and said, Surely the LORD'S anointed is before him.")
    v("[7] But the LORD said unto Samuel, Look not on his countenance, or on the height of his stature; because I have refused him: for the LORD seeth not as man seeth; for man looketh on the outward appearance, but the LORD looketh on the heart.")
    v("[11] And Samuel said unto Jesse, Are here all thy children? And he said, There remaineth yet the youngest, and, behold, he keepeth the sheep.")
    sp()

    ref("1 Samuel 17:28, 34-36")
    v("[28] And Eliab his eldest brother heard when he spake unto the men; and Eliab's anger was kindled against David, and he said... with whom hast thou left those few sheep in the wilderness?")
    v("[34] And David said unto Saul, Thy servant kept his father's sheep, and there came a lion, and a bear, and took a lamb out of the flock:")
    v("[35] And I went out after him, and smote him, and delivered it out of his mouth: and when he arose against me, I caught him by his beard, and smote him, and slew him.")
    v("[36] Thy servant slew both the lion and the bear: and this uncircumcised Philistine shall be as one of them...")
    sp()

    b([
        ("The teaching provides two additional Old Testament characters that illustrate the outcomes of divine testing: ", False, False),
        ("Eliab", True, False),
        (" and ", False, False),
        ("David", True, False),
        (". When the prophet Samuel came to Bethlehem to anoint a king in 1 Samuel 16, he looked on Eliab's impressive height and majestic countenance and assumed, 'Surely the Lord's anointed is before him.' But God said, 'Look not on his countenance, or on the height of his stature; because I have refused him.' The speaker highlights: ", False, False),
        ("'The Hebrew is like someone who failed the test.'", True, True),
        (" Eliab had external charisma and military stature, but inwardly his heart was bitter, contemptuous, and proud (as revealed when he attacked David in 1 Samuel 17:28). Eliab failed God's character test and was rejected.", False, False)
    ])

    b([
        ("David, however, passed the test of faithfulness in complete obscurity. When Samuel asked if all Jesse's sons were present, Jesse answered that one remained in the wilderness: ", False, False),
        ("'And, behold, he keepeth the sheep. That is what was known of him, till redemption: he is keeping the sheep!'", True, True),
        (" Even when Eliab mocked him at the battle, asking, 'With whom hast thou left those few sheep in the wilderness?', David had already placed them in good care. When brought before King Saul, David's credential was his pastoral vigilance: ", False, False),
        ("'Sheep! Said, I'm a faithful shepherd. When the bear and the lion came, your servant went after them. God gave me the power and I tore them into pieces. That means he didn't lose a sheep! So he mirrors the good shepherd. Faithful with the few. God walked through the sheep with David, then He led the nation, just like Moses, 40 years faithful with the sheep, then God gave him Israel.'", True, True),
        (" Faithfulness with the few sheep in the wilderness was the test that qualified David to rule the nation.", False, False)
    ])

    sp()
    add_section_h("CONCLUSION")

    b([
        ("In summary, God actively tests believers—not to orchestrate their failure, but as an engineer test-drives a newly manufactured vehicle to authenticate its structural integrity, endurance, and roadworthiness. Tests and faithfulness exist in an unbroken relationship: faithfulness cannot be verified in ease; it is forged and demonstrated when pressure is applied. God does not use men that He does not test; He tests you so that He can use you for more.", False, False)
    ])

    b([
        ("In the government of the local church, human leadership is commanded to prove workers before assigning responsibility (1 Tim 3:10, 13; 2 Tim 2:15). God permits fires in the desert so that those who cannot endure the wilderness do not cross over into the Promised Land to corrupt the flock. Untested leadership is an ecclesiastical hazard.", False, False)
    ])

    b([
        ("The teaching presents two clear outcomes to every test: We will either emerge like ", False, False),
        ("Joseph and David", True, False),
        ("—passing consecutive tests of slavery, false accusations, obscurity, and the defense of the few sheep without grumbling or bitterness, emerging with the character of Christ to rule nations; or we will end like ", False, False),
        ("Gehazi and Eliab", True, False),
        ("—relying on cosmetic appearance, harboring secret ambition or covetousness, failing the test of motives, and walking into spiritual rejection, disqualification, and leprosy. Every cell leader must welcome tests as God's instrument to build character and strength, remaining faithful at their duty post until they stand approved unto God.", False, False)
    ])

    # Save Markdown file
    md_content = "\n".join(md_lines)
    for p in md_paths:
        with open(p, "w", encoding="utf-8") as f:
            f.write(md_content)
        print(f"[+] Saved Markdown -> {p}")

    # Save Word document
    for p in docx_paths:
        doc.save(p)
        print(f"[+] Saved Word Document -> {p}")

if __name__ == "__main__":
    repo_root = Path(__file__).resolve().parent
    study_group_dir = repo_root / "Study group"

    target_docx = [
        repo_root / "Cell_Leaders_Conference_Faithfulness_in_Ministry_StudyGroup_Answers.docx",
        study_group_dir / "Cell_Leaders_Conference_Faithfulness_in_Ministry_StudyGroup_Answers.docx"
    ]

    target_md = [
        repo_root / "Cell_Leaders_Conference_Faithfulness_in_Ministry_StudyGroup_Answers.md",
        study_group_dir / "Cell_Leaders_Conference_Faithfulness_in_Ministry_StudyGroup_Answers.md"
    ]

    print("Building Faithfulness in Ministry Study Group Answers...")
    build_answers(target_docx, target_md)
    print("Done!")
