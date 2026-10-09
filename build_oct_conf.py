#!/usr/bin/env python3
"""
build_oct_conf.py
Generates Oct_Conf_Formatted.docx, Cell_Leaders_Conference_How_Leadership_Changes_You.docx,
and matching markdown versions.
"""

import os
from pathlib import Path
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH

def create_docx(output_path):
    doc = Document()

    # Page setup: Letter size, 1-inch margins (matching 12240 x 15840 dxa, 1440 dxa margin)
    for section in doc.sections:
        section.page_width = Inches(8.5)
        section.page_height = Inches(11.0)
        section.top_margin = Inches(1.0)
        section.bottom_margin = Inches(1.0)
        section.left_margin = Inches(1.0)
        section.right_margin = Inches(1.0)

    # Style defaults
    style = doc.styles['Normal']
    font = style.font
    font.name = 'Arial'
    font.size = Pt(12)
    font.color.rgb = RGBColor(0x11, 0x18, 0x27) # Dark neutral

    # Helper functions matching build_oct_conf.js
    def add_title(text):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(2) # 40 dxa
        run = p.add_run(text)
        run.bold = True
        run.font.name = 'Arial'
        run.font.size = Pt(18) # 36 half-pt

    def add_subtitle(text):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(10) # 200 dxa
        run = p.add_run(text)
        run.bold = True
        run.italic = True
        run.font.name = 'Arial'
        run.font.size = Pt(14) # 28 half-pt

    def h(text):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(14) # 280 dxa
        p.paragraph_format.space_after = Pt(4)  # 80 dxa
        run = p.add_run(text)
        run.bold = True
        run.font.name = 'Arial'
        run.font.size = Pt(12)
        return p

    def ref(text):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(8) # 160 dxa
        p.paragraph_format.space_after = Pt(2)  # 40 dxa
        run = p.add_run(text)
        run.bold = True
        run.italic = True
        run.font.name = 'Arial'
        run.font.size = Pt(12)
        return p

    def v(text):
        p = doc.add_paragraph()
        p.paragraph_format.left_indent = Inches(0.25) # 360 dxa
        p.paragraph_format.space_before = Pt(1) # 20 dxa
        p.paragraph_format.space_after = Pt(1)  # 20 dxa
        run = p.add_run(text)
        run.bold = True
        run.italic = True
        run.font.name = 'Arial'
        run.font.size = Pt(12)
        return p

    def sp():
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(3) # 60 dxa
        p.paragraph_format.space_after = Pt(3)  # 60 dxa
        return p

    def b(runs_data):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(4) # 80 dxa
        p.paragraph_format.space_after = Pt(4)  # 80 dxa
        p.paragraph_format.line_spacing = 1.15
        for r_text, r_bold, r_italic in runs_data:
            run = p.add_run(r_text)
            run.font.name = 'Arial'
            run.font.size = Pt(12)
            if r_bold:
                run.bold = True
            if r_italic:
                run.italic = True
        return p

    # --- Document Content ---
    add_title("Cell Leaders Conference")
    add_subtitle("How Leadership Changes You")

    # Section 1
    h("Leadership Is Service: The God Model vs the World Model")
    ref("Matthew 20:24-28")
    v("[24] And when the ten heard it, they were moved with indignation against the two brethren.")
    v("[25] But Jesus called them unto him, and said, Ye know that the princes of the Gentiles exercise dominion over them, and they that are great exercise authority upon them.")
    v("[26] But it shall not be so among you: but whosoever will be great among you, let him be your minister;")
    v("[27] And whosoever will be chief among you, let him be your servant:")
    v("[28] Even as the Son of man came not to be ministered unto, but to minister, and to give his life a ransom for many.")
    sp()
    b([("The idea some of us have about leadership needs to be renewed. In the secular and political world, leadership refers to somebody who has power, influence and authority - the police, the military, the apparatus of the state. That is how leadership is seen in the world system. But Jesus gives us God's idea. When James and John's mother came asking for the seats of honour, the ten were angry - and notice they were angry because they equally wanted it. Jesus said: whosoever will be great among you, let him be your minister. Whosoever will be chief, let him be your servant. The son of man came not to be ministered unto but to minister.", False, False)])
    sp()
    b([("Some people's idea of a minister is somebody who sits in front. That is not it. A minister is somebody who is prepared to give. Do not think that being in leadership is about perks of office. What comes with the office is what God has invested in the office - so that you can serve well. Without honor and respect for the vessel God is using, you cannot receive from it. But the purpose of that honor is not for the vessel to use it for personal ambition. No - it is so that they can serve the people well. Let your life command honor. Do not demand it. Let your service command it. Your speech, your generosity, your prayer life, your sacrifice - these are what command it. Honor is instructed in Scripture, yes. But honor is commanded by service, not demanded by title.", False, False)])
    sp()

    # Section 2
    h("Jesus as the Servant-Son: Acts 3 and Isaiah 42")
    ref("Acts 3:13")
    v("[13] The God of Abraham, and of Isaac, and of Jacob, the God of our fathers, hath glorified his Son Jesus.")
    sp()
    b([
        ("The word there for Son is ", False, False),
        ("pais", True, True),
        (" in the Greek - not the usual ", False, False),
        ("huios", True, True),
        (" or ", False, False),
        ("teknon", True, True),
        (". ", False, False),
        ("Pais", True, True),
        (" is used for a servant who is a son. A son-servant. A grown male servant who willingly yields himself for service. Acts 4:27 uses the same word for the holy child Jesus. And this word connects directly to Isaiah 42:1:", False, False)
    ])
    sp()
    ref("Isaiah 42:1")
    v("[1] Behold my servant, whom I uphold; mine elect, in whom my soul delighteth.")
    sp()
    b([
        ("That is the same ", False, False),
        ("pais", True, True),
        (" when the Septuagint renders it. A son-servant. He willingly yields his life. He is worthy of worship not because he demands it but because he gave himself. So you and I must see our service in that light. We command it. Our lives command it. Our generosity commands it. Our sacrifice commands it.", False, False)
    ])
    sp()
    ref("Philippians 2:5-9")
    v("[5] Let this mind be in you, which was also in Christ Jesus:")
    v("[6] Who, being in the form of God, thought it not robbery to be equal with God:")
    v("[7] But made himself of no reputation, and took upon him the form of a servant.")
    v("[9] Wherefore God also hath highly exalted him, and given him a name which is above every name.")
    sp()
    b([("What we often take from Philippians 2:5-9 is the authority of the believer. But look at the context. Paul was saying: let this mind be in you - the mind of one who, though he was God, became a servant. And God highly exalted him and gave him the name. He proved his status by service. Service proves our divinity. The life of God in us is a life of service. The life of the Spirit is a life of service. We are not more divine than ourselves. Service proves that we are of God.", False, False)])
    sp()

    # Section 3
    h("God Chooses Regular People: The Twelve")
    b([("Leadership is meant to change you. God always chooses normal human beings. And that is why when you begin to serve, you will have your weaknesses - many of them. You come from different backgrounds. You can be poor or affluent, a fisherman or an accountant. But the stream of service begins to change your life. Look at the kind of people Jesus chose, because it tells us something about how God works.", False, False)])
    sp()
    ref("John 1:38-42")
    v("[38] And they said unto him, Rabbi, where dwellest thou?")
    v("[40] One of the two which heard John speak, and followed him, was Andrew, Simon Peter's brother.")
    v("[41] He first findeth his own brother Simon, and saith unto him, We have found the Messias.")
    sp()
    b([("Andrew was the first person Jesus called - but Andrew was not the leader. He chose Peter. That Andrew brought. Andrew had the power of persuasion. He was able to convince others to come. But there are some who will bring people into the team and into the church - and that does not amount to being leaders. Who did Jesus choose to lead? Peter. Not because Peter was correct - Peter was a crude, forward, courageous, impulsive man who even had the guts to challenge Jesus. He talked before he thought. Rush. Yet Jesus chose him. And over time, you saw that leadership transformed him.", False, False)])
    sp()
    b([("In Acts 5 you can still see traces of the old Peter. In Acts 8, Acts 15. In Paul's confrontation of Peter in Galatians 2. And then you see the Peter of 2 Peter 3. Leadership transformed him. He moved from that impetuous fellow to the Peter who wrote with depth and pastoral care. God chose him the way he was - wrong. And he used that responsibility to change him.", False, False)])
    sp()
    b([("We had James and John - ambitious business partners who knew how to work influence. They went through their mother. Ambitious guys. Jesus chose them as young men. We had Nathaniel in John 1:46 who said: can anything good come out of Nazareth? An ethnic bias. He said what he thought without weighing his words. God chose him. Matthew was a tax collector - regarded as corrupt by his society, mixing with the wrong crowd. But he was detailed, kept accurate records, had good memory. That same skill shows up in how systematically he organized his Gospel. Jesus chose him. Thomas spoke in terms of reality. He said: let us go and die with him. He needed to see to believe. Jesus chose him. Simon the Zealot was a political activist - a political revolutionary with a terrible temper. Jesus chose him.", False, False)])
    sp()
    b([("And Jesus chose Judas Iscariot to handle the purse. He was not faithful with it - yet Jesus chose him. And that didn't become known until the end. Which means Judas kept the accounts throughout. Jesus chose everyday, regular people. Don't look for shining diamonds. God is going to use regular people. When he uses them - that is leadership changing them. There is no one whose prayer life became intense outside of people who started serving. Your prayer life becomes intense because you now have responsibility. Leadership transforms.", False, False)])
    sp()

    # Section 4
    h("You Were Chosen Imperfect: Don't Drop the Ball")
    b([("You were not chosen by God because you were perfect. You were chosen by God because you are imperfect - and he is going to use that responsibility to transform your life. Which means two things. First: I must not be proud when I fail. Whether I fail morally or fail in my duties, I should not say I am no longer doing this. I should say: now it is time to get it right. Because God gave me this task so that I can get better. When you drop it, it is because you are proud.", False, False)])
    sp()
    b([("The worst thing that can happen to a leader is to have people who are solely looking for fault. The moment you begin to focus on what you perceive as wrong in a leader, there are things bigger than your mouth. In Numbers 12, Moses took an Ethiopian wife and his siblings Aaron and Miriam said: has God not spoken through us also? Who does he think he is? Are we not all the same in Christ? You would expect God to say, I understand what you are saying. He did not say that. He said: were you not afraid to speak what you have said? God has a way he handles his servants. It is not by you.", False, False)])
    sp()
    b([("So don't get to the point of making it your mission to expose a leader. In Genesis 20, Abraham did not tell the truth about Sarah, and Abimelech got into trouble. God did not answer those kinds of questions. He just moved to fix the situation. He knows how to judge his own servants. And it is pride that makes a leader drop the ball - pride that says, I cannot serve God like this. He chose you in spite of you. He called you in spite of you. When you fail and falter, get back into it. That river, that stream of the Spirit, will keep developing you.", False, False)])
    sp()

    # Section 5
    h("Samson: Leadership Works Even in Imperfection")
    ref("Judges 13:24-25")
    v("[24] And the woman bare a son, and called his name Samson: and the child grew, and the LORD blessed him.")
    v("[25] And the Spirit of the LORD began to move him.")
    sp()
    b([("Samson was set apart from birth. His mother was given specific instruction - no wine or strong drink, and the child shall be a Nazarite from the womb to the day of his death. That instruction was not for everybody. It was given to the mother and then to him. There are things the Lord will say to you that are not to be shared.", False, False)])
    sp()
    b([("Interestingly, in Judges 14:1, Samson saw a woman of Timnath, a daughter of the Philistines, and said: go get me that wife. His parents said: there is no one among our own people? He said: get her for me. Look at verse 4: his father and mother did not know that it was of the Lord - that he sought an occasion against the Philistines, for at that time the Philistines had dominion over Israel. God used even the wrong thing. Not that God caused it - but God will use everything. Weaknesses, strength, imperfections. He is Lord over all of it.", False, False)])
    sp()
    b([("In Judges 15 he went in to a harlot. They surrounded him intending to kill him. He got up and defeated them all. In Judges 16, they surrounded him again - verse 2: lay wait for him all night and said, we will kill him in the morning. He arose at midnight and carried away the gates of the city. Again and again, the anointing functioned. Don't forget the instruction - he should not shave his head. That was the consecration. The consecration was separate from his moral life. The anointing and the immorality co-existed for a season because God is faithful. What eventually took his ministry away was not his immorality. It was foolish disclosure - lack of discretion. Delilah pressed him daily and his soul was vexed unto death, and he told her all that was in his heart. That is what took it away. Lack of discretion.", False, False)])
    sp()

    # Section 6
    h("Marriage and Ministry: You Don't Share Ministry")
    b([("Being married to a servant of God is not access to their calling. Marriage is meant to be complementary - to help us fulfill God's call. Not a leeway to cross the boundaries. What God says to you according to your role is specific. In Genesis 18, when God appeared to Abraham, he did not say, let me wait until Sarah comes from the market. God communicates with people according to their role. What Sarah needed to know concerning Isaac, God told her - in Genesis 18 when she laughed. What concerned Abraham's task of building a nation, God communicated that to Abraham alone. Sarah's role was to be his associate. She understood her role and held it.", False, False)])
    sp()
    b([("Delilah's approach is the model of manipulation in a close relationship: how can you say you love me and not tell me everything? We are one flesh. How can you keep secrets from me? Verse 15: how can you say you love me when your heart is not with me? She pressed him daily. His soul was vexed unto death. Manipulation. She weaponized love to extract what she had no right to know. And it took his ministry. There are people who function that way - the food is not sweet, the mood in the house is always heavy, everything is a pressure point - until the man or woman of God gives in.", False, False)])
    sp()
    b([("Look at Michal in 2 Samuel 6:20. David was dancing before the ark, worshipping God with all his strength. Michal looked out of the window and said: how glorious was the king of Israel today, who uncovered himself in the eyes of the handmaids of his servants as one of the vain fellows shamelessly uncovered himself. She used the language of shame to attack his worship. David's response was: it was before the Lord who chose me over your father. I will be yet more vile than thus. And 2 Samuel 6:23 records that she had no child to the day of her death. Use your mouth wisely in relation to your spouse's ministry.", False, False)])
    sp()
    b([("Look also at Job's wife - Job 2:9. She said: do you still retain your integrity? Curse God and die. The Bible calls it speaking foolishly. The same woman who celebrated his prosperity, when the pressure came, spoke foolishly. Don't be stupid with your lips. And do not get home and share details about the people in your care with your spouse or close friends. People come to you with what they have because of who you are. Don't make it domestic conversation. When God is speaking to you about your assignment, he will speak to your spouse concerning what is relevant to them. But the ministry's work is not shared. You are responsible for what you carry.", False, False)])
    sp()

    # Section 7
    h("Leadership Changes How You Think, Speak and Spend")
    b([("Leadership changes you. It changes how you think. Before you were a cell leader or pastor, there was a way you would talk about the church - meetings are too many, we don't need all of this. Now you are there. Your mouth will change. Your character is being transformed. You will learn patience - not just as fruit of the Spirit, but learned in the Spirit, by the Spirit. You learn how not to be in a hurry. You learn how not to be angry. God will bring somebody into your life who will change you. All your doctrine will be reformatted. Someone comes into your life for a purpose, and your fasting will change, your priorities will change.", False, False)])
    sp()
    b([("You know you are being changed when you look at yourself and know you have not done enough. You are burdened. You are in tears. You stay back after church when you could have left. You are looking around for who is not here. You are no longer focused on yourself - you are focused on the work. When you are negotiating your salary, something inside you knows there are responsibilities to the kingdom. You are changing. If you are not seeing it that way - then you are resisting the transformation.", False, False)])
    sp()
    b([("It is pride that makes you become self-conscious. I don't dress like a minister. I don't speak like a minister. I don't have a minister's car. You are going to lose the plot. God chose you the way you are. As you grow, he brings people into your life to mentor you, strengthen you, harness what is in you. The stream of leadership transforms you. I am not who I was over 30 years ago. I am not who I was 25 years ago. I am not who I was last year. And I am still leading. I am still changing. Leadership is a process, not a destination.", False, False)])
    sp()

    # Section 8
    h("Don't Give Up: Process and Patience")
    b([("One of the mistakes young leaders make is expecting perfection - from themselves or from others. When it doesn't work the way they planned, they throw in the towel. When you stop serving, you stop growing. You are running away from the very reason you were chosen. It is required in a steward that a man be found faithful - 1 Corinthians 4:2. God will use your errors to make you a better person. The mistake will become a ministry. What you felt you should not have done will one day be the thing you share with a younger person to guide them through the same territory.", False, False)])
    sp()
    b([("Jesus said to Peter: Simon, Satan has desired to sift you as wheat - but I have prayed for you, that your faith fail not. And when you are converted, strengthen the brethren. That is a positive leader. He did not say: since you are going to fail, I am removing you. He prayed for him. He prepared him for what was coming. And he called him to use his failure as a ministry to others. That is the model. Believe the best about your disciples. An elder who does not get hungry easily will have many children. God will bring raw human beings to you. Believe the best. Do not write people off in your heart when they fall.", False, False)])
    sp()
    b([("Those who give up easily are impatient and hate process. Information without process does not become wisdom. You can have money and still be an idiot because process makes you mature, not information alone. Some of you who are in this generation want to chop sharp sharp - quick results, quick growth, quick everything. When June passes and the money is not there, you go to another business. Relax. Process is not optional. Wisdom is when you sit with information and allow it to build you over time. That is what gives depth to ministry, to marriage, to career. Don't be impatient.", False, False)])
    sp()
    b([("Also: never let your hurts affect what you are doing. You will always get hurt - in business, in marriage, in ministry, in family. Never bring your hurts to the pulpit. Clean your eyes before you get there. Never allow the wounds of one relationship to poison the next one you are building. Learn to separate your pain from your assignment. You carry burdens, yet you lead others. You are imperfect, yet you are working on others. You are being perfected at the same time as you perfect others. That is the grace of God. Trust it.", False, False)])
    sp()

    # Section 9
    h("God Is Loyal to You: Don't Throw In the Towel")
    b([("No matter what men say about you, God is loyal to you. He's loyal to you. For the sake of his word, for the sake of his purpose and plan in the earth, he is loyal to you. So don't throw in the towel. When you fail in anything, it is an opportunity to deepen your leadership, to improve. It doesn't mean you are no longer who God wants to use. He chose you wrong. He called you when those things were already there. But he doesn't want them to remain there. Get back into it. That river, that stream of the Spirit, will keep developing you. It is pride that says, I can't serve God like that. He chose you in spite of you.", False, False)])
    sp()
    b([("You must function in the grace of God. Never see yourself as an entitled person in the ministry - it is the grace of God working in you. And because you know that God is using you in spite of yourself, let that same graciousness flow to those you lead. You know what God accommodated in you when he chose you. Apply the same standard to your disciples. Believe the best about people. He will bring them to you raw. What you are able to do with them is what will happen in the end.", False, False)])
    sp()
    b([("Leadership should change you. Transform you. It changes how you see, how you speak, how you spend, how you relate, how you worship. Say this: God is transforming me by leadership. My priorities are being reset. My cell leadership is changing me. My pastor is changing me. Those I am accountable to are changing me. I am being changed. I am changed. I will never stop being a leader. I will never get tired of doing it. I will never get bored of it. The Spirit of God is working on me in this season - and I am trusting him more for my life, for my relationships, for my ministry.", False, False)])

    doc.save(output_path)
    print(f"[+] Saved docx to: {output_path}")

def count_words(doc_path):
    doc = Document(doc_path)
    total_words = 0
    for p in doc.paragraphs:
        words = p.text.split()
        total_words += len(words)
    return total_words

if __name__ == '__main__':
    base_dir = Path(r"c:\Users\Stephycopy\OneDrive\Desktop\Theos\theos")
    out1 = base_dir / "Oct_Conf_Formatted.docx"
    out2 = base_dir / "Cell_Leaders_Conference_How_Leadership_Changes_You.docx"

    create_docx(out1)
    create_docx(out2)

    words1 = count_words(out1)
    print(f"Done. Words: {words1} / 13300 original")
