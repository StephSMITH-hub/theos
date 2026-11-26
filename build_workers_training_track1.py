# -*- coding: utf-8 -*-
"""
Master Builder for:
Workers in Training Convention 2026 – What to do with a new convert - Track 1
Speaker: Pastor Chris Segun Onayinka, Saints Community Church
Generates:
1. Workers in Training Convention 2026 – What to do with a new convert - Track 1 (Structured & Clean).docx (Root)
2. Workers in Training Convention 2026 – What to do with a new convert - Track 1 (Structured & Clean).txt (Root)
3. Workers in Training Convention Transcript/Raw/Workers in Training Convention 2026 – What to do with a new convert - Track 1_Raw.docx
4. Workers in Training Convention Transcript/Raw/Workers in Training Convention 2026 – What to do with a new convert - Track 1_Raw.txt
"""

import os
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH

def create_document():
    doc = docx.Document()
    
    # Page setup - 1.0 inch margins all around
    sections = doc.sections
    for section in sections:
        section.top_margin = Inches(1.0)
        section.bottom_margin = Inches(1.0)
        section.left_margin = Inches(1.0)
        section.right_margin = Inches(1.0)
        
    return doc

def build_content(doc, txt_lines):
    def add_title(text):
        p = doc.add_paragraph()
        p.paragraph_format.line_spacing = 1.15
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(2)
        run = p.add_run(text)
        run.font.name = "Lexend"
        run.font.size = Pt(18)
        run.font.bold = True
        run.font.color.rgb = RGBColor(15, 23, 42)
        txt_lines.append(text.upper() + "\n" + "=" * len(text))

    def add_meta(text):
        p = doc.add_paragraph()
        p.paragraph_format.line_spacing = 1.15
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(8)
        run = p.add_run(text)
        run.font.name = "Lexend"
        run.font.size = Pt(10)
        run.font.italic = True
        run.font.color.rgb = RGBColor(71, 85, 105)
        txt_lines.append(text + "\n")

    def add_h1(text):
        p = doc.add_paragraph()
        p.paragraph_format.line_spacing = 1.15
        p.paragraph_format.space_before = Pt(14)
        p.paragraph_format.space_after = Pt(4)
        p.paragraph_format.keep_with_next = True
        run = p.add_run(text)
        run.font.name = "Lexend"
        run.font.size = Pt(13)
        run.font.bold = True
        run.font.color.rgb = RGBColor(30, 58, 138)
        txt_lines.append("\n" + text.upper() + "\n" + "-" * len(text))

    def add_p(text):
        p = doc.add_paragraph()
        p.paragraph_format.line_spacing = 1.15
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(5)
        run = p.add_run(text)
        run.font.name = "Lexend"
        run.font.size = Pt(10)
        run.font.color.rgb = RGBColor(30, 41, 59)
        txt_lines.append(text)

    def add_ref(ref_text):
        p = doc.add_paragraph()
        p.paragraph_format.line_spacing = 1.15
        p.paragraph_format.space_before = Pt(6)
        p.paragraph_format.space_after = Pt(2)
        p.paragraph_format.keep_with_next = True
        run = p.add_run(ref_text)
        run.font.name = "Lexend"
        run.font.size = Pt(10)
        run.font.bold = True
        run.font.color.rgb = RGBColor(15, 23, 42)
        txt_lines.append("\n" + ref_text)

    def add_verse(verse_text):
        p = doc.add_paragraph()
        p.paragraph_format.line_spacing = 1.15
        p.paragraph_format.left_indent = Inches(0.25)
        p.paragraph_format.space_before = Pt(1)
        p.paragraph_format.space_after = Pt(3)
        run = p.add_run(verse_text)
        run.font.name = "Lexend"
        run.font.size = Pt(9.5)
        run.font.bold = True
        run.font.color.rgb = RGBColor(30, 41, 59)
        txt_lines.append("    " + verse_text)

    # Document Header
    add_title("Workers in Training Convention 2026")
    add_meta("Teaching: What to do with a new convert - Track 1 | Speaker: Pastor Chris Segun Onayinka | Ministry: Saints Community Church")

    # Section 1
    add_h1("1. Introduction: The Nature and Origin of Workers in Church History")
    add_p("A workers' convention like this, or Workers in Training Convention as we call it in Saints Community Church, is a special meeting.")
    add_p("The individuals here are selected, not indiscriminately, but they are people who have shown willingness to serve and willingness to be available to serve other saints.")
    add_p("So we call them workers.")
    add_p("Naturally, you won't see that phrase in scripture.")
    add_p("I guess one of the things we have to include in our book, Signs of the Times, is workers and non-workers, almost looking like employed and unemployed.")
    add_p("You won't find that phrase in scripture.")
    add_p("Over time, we have seen that based on historical developments, some things are historical developments.")
    add_p("That is, you won't necessarily see them in scripture, but historically they have come up; that is, they have been developed.")
    add_p("Sometimes it has become a doctrine, and we must be careful not to make it such. That is why we are doing our best not to adopt it.")
    add_p("You will notice in Matthew 9:36, Matthew records where Jesus had gone round the cities. He repeats Matthew 4:23 in Matthew 9:35; it is like word-for-word verbiage.")
    add_p("Then he says Jesus looked on the multitude, that is, having crowds of people.")
    add_p("They were scattered abroad, and he had compassion on them.")
    add_p("Why did he have compassion on them?")
    add_p("Don't forget he had taught them the word. He had preached the gospel of the kingdom to them. He had healed all manner of sickness and diseases. He had also cleansed the lepers.")
    add_p("Yet he said he had compassion on them because they were like sheep that had no shepherd.")
    add_p("That is interesting, because one would have thought a church service is so glorious when people hear the word, when their sicknesses are cured and healed, and when they receive miracles.")
    add_p("That should pass for a glorious convention. By many parameters today, that works; that is a good meeting.")
    add_p("Yet he still had compassion.")
    add_p("At the end of a convention, after people have said, 'See what the Lord has done; he healed my body, he cast out demons from my body,' and everyone felt that was good, he still had compassion on them, for they were like sheep having no shepherd.")

    add_ref("Matthew 9:35–38 (KJV)")
    add_verse("[35] And Jesus went about all the cities and villages, teaching in their synagogues, and preaching the gospel of the kingdom, and healing every sickness and every disease among the people.")
    add_verse("[36] But when he saw the multitudes, he was moved with compassion on them, because they fainted, and were scattered abroad, as sheep having no shepherd.")
    add_verse("[37] Then saith he unto his disciples, The harvest truly is plenteous, but the labourers are few;")
    add_verse("[38] Pray ye therefore the Lord of the harvest, that he will send forth labourers into his harvest.")

    # Section 2
    add_h1("2. Sheep Having No Shepherd: Direction, Destination, and the Abrahamic Covenant")
    add_p("If you look at the background of that phrase, because the Old Testament provides a background for such information, in Numbers 27 it was when God had repeatedly told Moses that he would not be leading Israel to the Promised Land.")
    add_p("Moses beckoned on God and said, 'You cannot allow your people to be as sheep having no shepherd.'")
    add_p("You need a man, Moses says in Numbers 27:16–18, who will lead them in and out, who will guide them.")
    add_p("You need a man, or rather a woman, because sometimes when you say that, in people's heads it is male.")
    add_p("I know a major denomination in America just voted two days ago not to have women pastors. They voted that a woman is not meant to lead a congregation.")
    add_p("It was a vote, and there was no transmission of results dispute. They just chose and said a woman cannot lead.")
    add_p("All the women here who are cell leaders and church pastors, they just disqualified you.")
    add_p("They voted, and it was overwhelming, though some people argued against it.")
    add_p("I keep telling people, imagine if Jesus had chosen a woman as part of the twelve. This is 2026, and people are still arguing whether a woman can lead a congregation.")
    add_p("In that nation, they struggle to have a woman as their president, while they are at the forefront of women's rights.")
    add_p("Two women have contested in the last ten years, and they lost. The first one was white, the second one was black; they didn't accept it.")
    add_p("For a lot of people, even some of us here, God is helping us, but we struggle to see a woman give us instructions.")
    add_p("When people see a woman rise in her profession, they think something must have happened. Someone was telling me, 'Do you think that the lady who rose to be the managing director of one of the biggest banks in Africa just went there like that?'")
    add_p("This was a classmate. I expected that at that age he should have seen enough not to make that kind of statement, but he did.")
    add_p("I asked him a question: 'In your university—he attended OAU, by the way—who was the best student in your class?'")
    add_p("He struggled to answer. He knew the answer. He said, 'That's not the issue.' I said, 'No, answer that question.'")
    add_p("He said, 'Well, it was one girl like that.' She studied architecture and was the best student.")
    add_p("I said, 'How come? She had the ability to beat you.'")
    add_p("The point was, you can imagine someone in 2026 still thinking that way. How will they raise their daughters?")
    add_p("Some people raise their daughters just to be wives, or to be second hand.")
    add_p("When we keep repeating things, it is so that it will be settled in the minds of people. Repetition is the key to lasting conviction; you must hear it again and again.")

    add_ref("Numbers 27:16–18 (KJV)")
    add_verse("[16] Let the LORD, the God of the spirits of all flesh, set a man over the congregation,")
    add_verse("[17] Which may go out before them, and which may go in before them, and which may lead them out, and which may bring them in; that the congregation of the LORD be not as sheep which have no shepherd.")
    add_verse("[18] And the LORD said unto Moses, Take thee Joshua the son of Nun, a man in whom is the spirit, and lay thine hand upon him;")

    add_p("Moses said in Numbers 27:17, a man. That is what got me there. People ask, 'Why did God choose Moses and not Miriam? Why did Jesus choose men as the twelve apostles?' Those are the arguments.")
    add_p("Joshua was chosen for what purpose? I want you to take this very clearly: it was to lead the people to the Promised Land.")
    add_p("What Moses requested was not just somebody who would just be leading. It is good to know the role.")
    add_p("That phrase was used again in the book of Kings for the king of Israel, when Micaiah said they are all scattered abroad like sheep having no shepherd.")
    add_p("The king of Israel, whose purpose was to govern a people in the Promised Land.")

    add_ref("1 Kings 22:17 (KJV)")
    add_verse("[17] And he said, I saw all Israel scattered upon the hills, as sheep that have not a shepherd: and the LORD said, These have no master: let them return every man to his house in peace.")

    add_p("At that point they were in the Promised Land. What was the idea of the Promised Land?")
    add_p("It was a place they would build to serve God amongst all the nations in the earth.")
    add_p("The role and the job is not to occupy space and just be leading people. You must have direction.")
    add_p("Sheep having no shepherd is not just bearing the title Pastor. You must know the direction.")
    add_p("At the last cell leaders' conference, we taught on the key indicators of growth. When you are employed, you must know the vision.")
    add_p("You must know exactly what the Lord Jesus Christ wants the church to do. I call it Key Performance Indicators.")
    add_p("When you are employed anywhere, you know what the company wants to do. In the body of Christ, you must know exactly what it is, because some things can be primary and you make them secondary, or secondary and you make them primary.")
    add_p("What do you mean by sheep having no shepherd? The job of Moses was to lead them to the Promised Land.")
    add_p("There was a covenant God gave Abraham in Genesis 12:2–3, Genesis 15:3–6, and Genesis 17:4–7 that he renewed in Isaac and renewed in Jacob.")
    add_p("God is telling Moses, 'I have a covenant for all nations of the earth. I have a covenant for the whole world, and I am using you as my covenant partner to get the job done.'")
    add_p("Sheep having no shepherd is connected to the covenant God gave to Abraham, and Moses is meant to fulfill it. Moses is not meant to reinvent the wheel.")
    add_p("There are clear-cut indicators as to what is to be done. There is a particular land, a particular place, and in Abraham shall all the nations of the earth be blessed.")
    add_p("That land will now have a global vision for the whole world.")
    add_p("It carries more than just leading people every Tuesday, teaching them the word of God, exposing scriptures in a deep way, and having them enthralled and excited at your insight.")
    add_p("It is to lead them to God's promised land: where is the destination? Where is the Lord taking us to? What is his plan?")
    add_p("The world does not set the agenda for the church. The news does not do it.")
    add_p("It is the Spirit of God, the written word, and by the Spirit expressing that to us, we have the agenda written ahead of us.")

    add_ref("Genesis 12:2–3 (KJV)")
    add_verse("[2] And I will make of thee a great nation, and I will bless thee, and make thy name great; and thou shalt be a blessing:")
    add_verse("[3] And I will bless them that bless thee, and curse him that curseth thee: and in thee shall all families of the earth be blessed.")

    add_ref("Genesis 15:5–6 (KJV)")
    add_verse("[5] And he brought him forth abroad, and said, Look now toward heaven, and tell the stars, if thou be able to number them: and he said unto him, So shall thy seed be.")
    add_verse("[6] And he believed in the LORD; and he counted it to him for righteousness.")

    add_ref("Genesis 17:4–7 (KJV)")
    add_verse("[4] As for me, behold, my covenant is with thee, and thou shalt be a father of many nations.")
    add_verse("[5] Neither shall thy name any more be called Abram, but thy name shall be Abraham; for a father of many nations have I made thee.")
    add_verse("[6] And I will make thee exceeding fruitful, and I will make nations of thee, and kings shall come out of thee.")
    add_verse("[7] And I will establish my covenant between me and thee and thy seed after thee in their generations for an everlasting covenant, to be a God unto thee, and to thy seed after thee.")

    # Section 3
    add_h1("3. The Plenteous Harvest and Laborers: Conviction Born of Relevant Information")
    add_p("Back to Matthew 9: Jesus had compassion on people who had been healed, compassion on people who had been set free from bondages, and compassion on people who had been taught the word.")
    add_p("Why? Because they were like sheep having no shepherd.")
    add_p("He told them, 'The harvest truly is plenteous, but the laborers are few.'")
    add_p("What are the laborers there? He is referring to shepherds.")
    add_p("'Pray ye therefore the Lord of the harvest, that he will send forth laborers into his harvest.'")
    add_p("He is not just talking about evangelism, which is part of it, but laborers into his harvest.")
    add_p("Then in chapter 10, he chooses twelve people whom we call the apostles, and he sends them away.")

    add_ref("Matthew 10:1–2 (KJV)")
    add_verse("[1] And when he had called unto him his twelve disciples, he gave them power against unclean spirits, to cast them out, and to heal all manner of sickness and all manner of disease.")
    add_verse("[2] Now the names of the twelve apostles are these; The first, Simon, who is called Peter, and Andrew his brother; James the son of Zebedee, and John his brother;")

    add_p("I want to read a quote that I repeatedly share every year on my Facebook page by Pastor W.F. Kumuyi of the Deeper Life Christian Ministry.")
    add_p("I consider this quote a very powerful one because it summarizes many things in God's word.")
    add_p("He said: 'Nothing of eternal and rewardable value can be done without conviction.'")
    add_p("'Nothing of eternal and rewardable value can be done without conviction. Nothing.'")
    add_p("Conviction stands as the difference between the man who carries out the task to the end, and the man who jumps out of the ship in the middle of the road. He is not convinced enough.")
    add_p("He is not convinced that this is what is worth living for.")
    add_p("In teaching you on consecration, one of the forces we said was conviction.")
    add_p("Conviction is based on available and relevant information. If you are not hearing the right things, you cannot have the right conviction.")
    add_p("You must hear the right things, and you must hear it repeatedly so that you can build a conviction.")
    add_p("Conviction is a product of available and relevant information.")
    add_p("For example, if a local church focuses on politics and how the church is responsible for the governance of a country, and that Christians should aspire for political office—they have doctrines called the Seven Mountains where they say, 'Go and possess this mountain, possess that mountain'—that kind of person cannot be found in vigils praying for a long time.")
    add_p("He won't see a reason why he needs to go for evangelism regularly. He will say, 'All I need basically is to ensure that I hone my skills properly, or I know how to manipulate the perception of people so that I can win their heart like Absalom and win elections.'")
    add_p("He won't see a reason why he has to live right, because what he needs to do is to be adept with the skillfulness required in the industries projected as fulfilling God's plan.")
    add_p("Nothing in the plan of God says you must not do that, but eternal and rewardable rewards require conviction.")
    add_p("You need that information that will boost your actions in spite of challenges.")
    add_p("Do not compare yourself with other believers or your church with other churches; compare yourself with the word.")
    add_p("When there is no information available, it is not regular.")
    add_p("When I got saved and received the gospel under the Youth Foundation Incorporated Mission, they didn't give me room to be called a new convert.")
    add_p("The week afterwards, I went for Bible study, prayer meeting, and ministers' meeting.")
    add_p("A week afterwards, we had prayer groups. On Sunday and Monday, some people would come together and have a vigil, and you joined that group.")
    add_p("Monday and Tuesday, another group; Wednesday, everybody together; Thursday and Friday, another group for two days of vigil; then Saturday and Sunday.")
    add_p("I joined all the groups until I felt sick, because I was doing vigils every day. I am talking about the first time I heard the gospel.")
    add_p("All along, there had been no room to hear about the Christian faith differently from laying your life down to serve Jesus.")
    add_p("Nobody ever told me that I needed to possess one mountain while possessing another mountain; I didn't have access to that.")
    add_p("When you have that foundation, it will build a lasting conviction, because from a child you have known the holy scriptures, as Paul told Timothy in 2 Timothy 3:15, which made him wise unto salvation.")
    add_p("Proverbs 22:6 says, 'Train up a child in the way he should go: and when he is old, he will not depart from it.'")

    add_ref("Proverbs 22:6 (KJV)")
    add_verse("[6] Train up a child in the way he should go: and when he is old, he will not depart from it.")

    add_ref("2 Timothy 3:15 (KJV)")
    add_verse("[15] And that from a child thou hast known the holy scriptures, which are able to make thee wise unto salvation through faith which is in Christ Jesus.")

    add_p("Is it possible to have people who heard that and are not active for the Lord? Yes, they stopped hearing it.")
    add_p("When you stop hearing it, you pay attention to something else, and it will look like you had an option. You chose the option, but it was not God who gave you the option.")
    add_p("To have a lasting conviction, to work with eternal reward in view, you need to have a lasting conviction, and that conviction is a product of repetition.")
    add_p("We are a product of what we hear regularly. Everybody is.")
    add_p("You are a product of what you expose your mind to very regularly.")

    # Section 4
    add_h1("4. The Great Commission: Making Disciples of All Nations")
    add_p("We are going to look at what to do with a new convert or a new member: what do you do? You make disciples.")
    add_p("In Matthew 28:18, Jesus said, 'All authority is given unto me in heaven and on earth. Go ye therefore, and teach all nations.'")
    add_p("The word teach there means to make disciples. Matheteuo: to make disciples of all nations.")
    add_p("Baptizing them in the name of the Father, and of the Son, and of the Holy Ghost.")
    add_p("Why does the King James Version use teach? Because you cannot have disciples without teaching.")
    add_p("Verse 20 shows you the emphasis: 'Teaching them to observe all things whatsoever I have commanded you.'")
    add_p("Look at the keyword: 'whatsoever,' or 'all things whatsoever.'")
    add_p("If you put that together with 2 Timothy 2:2, 'the things that thou hast heard of me among many witnesses, the same commit thou to faithful men.'")
    add_p("He says the things, all things whatsoever.")
    add_p("Those two texts mean that you are not going to select what to pass across and what not to pass across.")
    add_p("All things whatsoever you heard of me.")
    add_p("You are not meant to invent it. You are not meant to sit down and think, 'What will work?' No.")
    add_p("You are not meant to say, 'What will people like?' No.")
    add_p("He said, 'Teaching them to observe all things whatsoever I have commanded you: and, lo, I am with you alway, even unto the end of the world.'")
    add_p("When we say, 'What to do to a new convert, what to do to a new member,' we must look at what Jesus did, and not just Jesus, but what the early apostles of Christ did, who laid for us the foundation.")

    add_ref("Matthew 28:18–20 (KJV)")
    add_verse("[18] And Jesus came and spake unto them, saying, All power is given unto me in heaven and in earth.")
    add_verse("[19] Go ye therefore, and teach all nations, baptizing them in the name of the Father, and of the Son, and of the Holy Ghost:")
    add_verse("[20] Teaching them to observe all things whatsoever I have commanded you: and, lo, I am with you alway, even unto the end of the world. Amen.")

    add_ref("2 Timothy 2:2 (KJV)")
    add_verse("[2] And the things that thou hast heard of me among many witnesses, the same commit thou to faithful men, who shall be able to teach others also.")

    # Section 5
    add_h1("5. The Calling of the First Disciples: Following is Deliberate and Sensitive")
    add_p("When he says to make disciples, how did Jesus make disciples?")
    add_p("What do you do to a new member or a new convert? You make disciples.")
    add_p("That short word, which requires explanation, must always be in your mind: to make disciples.")
    add_p("People can drop that word very easily and say, 'He's my disciple, she's my disciple,' and they may not even know what they are talking about.")
    add_p("Have it in your mind: to make disciples. What do you mean by making disciples? Notice that word: to make.")

    add_ref("Matthew 4:17–22 (KJV)")
    add_verse("[17] From that time Jesus began to preach, and to say, Repent: for the kingdom of heaven is at hand.")
    add_verse("[18] And Jesus, walking by the sea of Galilee, saw two brethren, Simon called Peter, and Andrew his brother, casting a net into the sea: for they were fishers.")
    add_verse("[19] And he saith unto them, Follow me, and I will make you fishers of men.")
    add_verse("[20] And they straightway left their nets, and followed him.")
    add_verse("[21] And going on from thence, he saw other two brethren, James the son of Zebedee, and John his brother, in a ship with Zebedee their father, mending their nets; and he called them.")
    add_verse("[22] And they immediately left the ship and their father, and followed him.")

    add_p("Don't miss the crucial details: they left and followed him.")
    add_p("This is a very sensitive work, because he has asked them to follow him.")
    add_p("When you ask people to follow you, you must be worthy of following.")
    add_p("You must know the seriousness of what it means for another human being to follow you.")
    add_p("You must know the sensitive nature, because exactly what he told them to do in Matthew 28:19–20 is what he did with them: 'Follow me, and I will make you.'")
    add_p("'I will make you.' That is what he said to Peter: 'I will make you.'")
    add_p("It is a sensitive work; you must ensure that you are worth following.")

    add_ref("Mark 1:16–20 (KJV)")
    add_verse("[16] Now as he walked by the sea of Galilee, he saw Simon and Andrew his brother casting a net into the sea: for they were fishers.")
    add_verse("[17] And Jesus said unto them, Come ye after me, and I will make you to become fishers of men.")
    add_verse("[18] And straightway they forsook their nets, and followed him.")
    add_verse("[19] And when he had gone a little further thence, he saw James the son of Zebedee, and John his brother, who also were in the ship mending their nets.")
    add_verse("[20] And straightway he called them: and they left their father Zebedee in the ship with the hired servants, and went after him.")

    add_ref("Luke 5:10–11 (KJV)")
    add_verse("[10] And so was also James, and John, the sons of Zebedee, which were partners with Simon. And Jesus said unto Simon, Fear not; from henceforth thou shalt catch men.")
    add_verse("[11] And when they had brought their ships to land, they forsook all, and followed him.")

    add_p("We have three accounts that show us how they became his disciples.")
    add_p("They had to make decisions based on his direction.")
    add_p("It therefore means that the one calling people to follow him or her must be worth following.")
    add_p("If you are that person, it therefore means you need to realize the sensitive nature of what you are doing.")
    add_p("It means therefore that your conduct matters. It matters a lot: the way you live your life.")
    add_p("Your decisions also matter: the kind of decisions you take. You cannot say, 'Well, it's my life, I can live it the way I want.' No, you can't say that, because you just told someone to follow you.")
    add_p("If you are going to tell someone to follow you, then they must be in your consideration.")
    add_p("They must be in your decision making.")
    add_p("It is a sensitive work. You don't use 'He is my disciple' in a loose way.")
    add_p("Your conduct, your decisions, your actions, and how you correct yourself in your inactions or misactions is vital.")
    add_p("It is not possible for someone to just walk up to another person by the riverside and say, 'Hello, good morning, follow me,' and they drop their net and follow.")
    add_p("You would call the police. That would be charm, or hypnotism, or something criminal.")
    add_p("It is not possible for someone to just come and say, 'You are fishermen, don't worry, you no longer fish fish, you fish men, let's go.' No.")
    add_p("The reason why he was able to say that to them was because they had known him.")
    add_p("They had known him before then; they weren't seeing him for the first time.")
    add_p("How did they know Jesus? Look at Acts of the Apostles chapter 1.")

    add_ref("Acts 1:21–22 (KJV)")
    add_verse("[21] Wherefore of these men which have companied with us all the time that the Lord Jesus went in and out among us,")
    add_verse("[22] Beginning from the baptism of John, unto that same day that he was taken up from us, must one be ordained to be a witness with us of his resurrection.")

    add_p("Beginning where? From the baptism of John. They knew Jesus through John's baptism.")
    add_p("From John's baptism, they had been following him.")
    add_p("Look at John 1:35. John stood, after John's baptism, and John in all four accounts—Matthew 3:11–12, Luke 3:16–17, Mark 1:8, and John 1:26–34—expressly described who Jesus was.")
    add_p("He had been preaching about Jesus, informing them well about Jesus.")
    add_p("The only issue was that he had not pointed out the person directly.")
    add_p("He was the voice crying in the wilderness, as Isaiah 40:3 and Mark 1:2–3 say: 'Prepare ye the way of the Lord, make his paths straight.'")
    add_p("He said, 'There standeth one among you whom ye know not; he it is who coming after me is preferred before me.'")
    add_p("He had been preaching about Jesus, describing Jesus and his ministry, including the baptism with the Holy Spirit.")
    add_p("Eventually, the Father spoke: 'This is my beloved Son, in whom I am well pleased.'")

    add_ref("John 1:35–37 (KJV)")
    add_verse("[35] Again the next day after John stood, and two of his disciples;")
    add_verse("[36] And looking upon Jesus as he walked, he saith, Behold the Lamb of God!")
    add_verse("[37] And the two disciples heard him speak, and they followed Jesus.")

    add_p("I have heard people mischievously use this to say that you could have somebody else pastor you at some point, and then someone else takes over.")
    add_p("What you are missing there is that John had said, 'I am not the Christ; I am only a voice.'")
    add_p("John's message was for people to follow Jesus. That is the distinction. You don't use that to steal people's members.")
    add_p("John's message was clear: Jesus was his focus. John's work was to make people follow Jesus.")
    add_p("The moment he pointed him out—'This is the Lamb of God I have been speaking about'—they followed Jesus.")
    add_p("They knew of Jesus already from John. They had heard about him; it was John's message.")
    add_p("They had listened to John about Jesus, and John introduced them to him.")
    add_p("Afterwards, they followed Jesus to preach. He was preaching the gospel of the kingdom, and they heard him preach.")
    add_p("There are instances where a person becomes your disciple after a while.")
    add_p("They had heard him preach and teach, they knew his mission, and they knew his work, even though they didn't have absolute information yet.")
    add_p("When they had heard about Jesus, he came to them as someone they already knew.")
    add_p("It wasn't just out of the blues that somebody came and said, 'Follow me.' The 'me' had been explained to them from John and from Jesus's teachings by himself.")
    add_p("He didn't make them disciples out of the blues.")

    # Section 6
    add_h1("6. Minister-to-Disciple Relationships in Paul's Ministry")
    add_p("You must know the basis of relationships. There are people that you just bless; you are a blessing to their Christian lives, and God uses you to bless them.")
    add_p("That is very different from this. These ones are not listening to teaching alone; they are following him.")
    add_p("You must know that difference.")
    add_p("There are different ways Paul speaks to churches.")
    add_p("Paul had Barnabas. Barnabas was someone they worked together with.")
    add_p("In Acts 9:27, many people did not know that Paul was a Christian; they didn't believe he was following Jesus.")
    add_p("Barnabas took him and brought him to the apostles, declaring how he had seen the Lord in the way.")
    add_p("In Acts 11:25–26, Barnabas went to Tarsus to seek Saul, and brought him to Antioch. Barnabas was a ministry associate, working together.")
    add_p("In Acts 13:2, the Spirit said, 'Separate me Barnabas and Saul for the work.' They were known together, and Paul recognized that in 1 Corinthians 9:6.")
    add_p("In Acts 15:39–40, Paul and Barnabas had a sharp contention over Mark, and they parted asunder. Barnabas took Mark and sailed to Cyprus.")
    add_p("Paul chose Silas, being recommended by the brethren unto the grace of God.")
    add_p("How did he choose Silas? The church recommended Silas to follow him. That is another kind of relationship.")

    add_ref("Acts 15:39–40 (KJV)")
    add_verse("[39] And the contention was so sharp between them, that they departed asunder one from the other: and so Barnabas took Mark, and sailed unto Cyprus;")
    add_verse("[40] And Paul chose Silas, and departed, being recommended by the brethren unto the grace of God.")

    add_p("In Acts 16, you have another person called Timothy.")
    add_p("Acts 16:1 says a certain disciple was there named Timotheus, well reported of by the brethren at Lystra and Iconium. Paul took him.")
    add_p("You have Timothy and you have Silas following him.")
    add_p("Paul will now call Timothy his beloved son.")
    add_p("He told Timothy in 2 Timothy 3:15, 'From a child thou hast known the holy scriptures.'")
    add_p("You have three distinct relationships: Barnabas as colleague, Silas as church-recommended co-laborer, and Timothy as son and disciple.")
    add_p("Then you have the churches that Paul founded: Lystra and Derbe in Acts 14, Macedonia and Philippi in Acts 16, Thessalonica in Acts 17, Corinth in Acts 18, and Ephesus in Acts 19.")
    add_p("In Acts 19:1, he found certain disciples at Ephesus who had only known John's baptism, and he instructed them.")
    add_p("Thessalonica, Philippi, Galatia, Corinth: all of them have letters that Paul wrote.")

    add_ref("Acts 16:1–3 (KJV)")
    add_verse("[1] Then came he to Derbe and Lystra: and, behold, a certain disciple was there, named Timotheus, the son of a certain woman, which was a Jewess, and believed; but his father was a Greek:")
    add_verse("[2] Which was well reported of by the brethren that were at Lystra and Iconium.")
    add_verse("[3] Him would Paul have to go forth with him; and took and circumcised him because of the Jews which were in those quarters: for they knew all that his father was a Greek.")

    # Section 7
    add_h1("7. Pastoral Authority, Transfer-Market Churchianity, and Social Media Immaturity")
    add_p("Then you have Rome. We have no record in the book of Acts about Rome, except where he went to after he was incarcerated.")
    add_p("Look at what he says to the church in Rome in Romans 15:18–20.")

    add_ref("Romans 15:18–20 (KJV)")
    add_verse("[18] For I will not dare to speak of any of those things which Christ hath not wrought by me, to make the Gentiles obedient, by word and deed,")
    add_verse("[19] Through mighty signs and wonders, by the power of the Spirit of God; so that from Jerusalem, and round about unto Illyricum, I have fully preached the gospel of Christ.")
    add_verse("[20] Yea, so have I strived to preach the gospel, not where Christ was named, lest I should build upon another man's foundation:")

    add_p("The church in Rome was different. He said, 'Lest I should build on another man's foundation.'")
    add_p("Sometimes what we do in the church world is transfer season. The football season has a transfer market where clubs buy players.")
    add_p("What we do sometimes in the pentecostal world is transfer season: one church donating new members to another church.")
    add_p("You now have what is called churchianity, or trying to woo fellow Christians to our church.")
    add_p("How do people do that? They ensure they find a fault in that church, and criticize what they preach.")
    add_p("Not that you love the brother or sister to know Jesus, but you criticize what they preach.")
    add_p("A minister in Nigeria spoke to me, really disturbed about his local church. I told him, 'That local church is God's church. We are not trying to woo people all over the churches.'")
    add_p("One thing that does is that you will never know what to do with a new convert. You won't know how to raise people or build lives; you will only know how to woo people.")
    add_p("I saw on a pastor's WhatsApp status—I am shocked at what a pastor can do, but these days you don't know who is Gen Z anymore. I think Gen Z is a philosophy, not age.")
    add_p("I see adults, even in church, come on WhatsApp and post emotional rants: 'Yeah, someone wants to kill me here o!' We don't even know who you are talking about.")
    add_p("They put how they are feeling on their status. Some criticize others or fight with friends on status.")
    add_p("Apart from spiritual training, some Christians need a mental check. If someone comes to you talking about someone you don't know, you would think the person is crazy.")
    add_p("You don't tell the world what is not their business. Who goes outside their home and shouts, 'Come and see what my wife said!' People will think you are mad.")
    add_p("Except your life is in danger, you don't speak to strangers about private matters.")
    add_p("A pastor put on his WhatsApp status: 'If you think you can keep your members from listening to us, you are lying; this gospel jumps fence, it is jumping into your bedroom!'")
    add_p("I called him and asked, 'Who are you talking to? Who are you talking about?'")
    add_p("He said a pastor in his city told his members that he was a heretic and shouldn't be listened to.")
    add_p("I asked him, 'Are you a heretic?' He said no. I said, 'Good. That is your responsibility: to ensure you are not a heretic.'")
    add_p("'It is his responsibility to warn people from heretics, and you have a duty not to be one. That's all.'")
    add_p("The nicest name I have been called is heretic; that has biblical explanation.")
    add_p("Someone claimed they saw me and my friend with calabash dancing round because we are in a cult, describing a Holy Ghost meeting in 1995.")
    add_p("Years later, that same person bought forty or fifty of our books.")
    add_p("In 1998, a man went on the pulpit and told everybody in church not to attend our camp meeting, claiming we were homosexuals and heretics preaching that there is no heaven and Jesus is not coming.")
    add_p("In 2016 or 2017, the same person went to another country and preached our message.")
    add_p("You must ensure you are not a heretic.")
    add_p("I asked the pastor, 'How many people are in that city?' He said about two million. 'How many members does that church have?' He said about one thousand.")
    add_p("I said, 'You still have 1,999,000 people. Why are you particular about that one thousand? That means you are not ready to reach people; you want to steal people.'")
    add_p("Focus on the people who are yet to believe the gospel.")

    add_ref("Romans 14:1–4 (KJV)")
    add_verse("[1] Him that is weak in the faith receive ye, but not to doubtful disputations.")
    add_verse("[2] For one believeth that he may eat all things: another, who is weak, eateth herbs.")
    add_verse("[3] Let not him that eateth despise him that eateth not; and let not him which eateth not judge him that eateth: for God hath received him.")
    add_verse("[4] Who art thou that judgest another man's servant? to his own master he standeth or falleth. Yea, he shall be holden up: for God is able to make him stand.")

    add_p("In Romans 14, Paul speaks about food and days with great care: 'Who art thou that judgest another man's servant?'")
    add_p("Why does he talk like that? Because it is Rome.")
    add_p("When he gets to Galatians, he doesn't say that; he says, 'You are fallen from grace. My little children, of whom I travail in birth again until Christ be formed in you.'")
    add_p("In 1 Corinthians 4:14–15, he says expressly: 'I write not these things to shame you, but as my beloved sons I warn you. For though ye have ten thousand instructors in Christ, yet have ye not many fathers: for in Christ Jesus I have begotten you through the gospel.'")
    add_p("He doesn't say that in Rome, because his relationship in Rome is very different from Corinth and Galatia.")

    add_ref("Galatians 4:19 (KJV)")
    add_verse("[19] My little children, of whom I travail in birth again until Christ be formed in you,")

    add_ref("1 Corinthians 4:14–17 (KJV)")
    add_verse("[14] I write not these things to shame you, but as my beloved sons I warn you.")
    add_verse("[15] For though ye have ten thousand instructors in Christ, yet have ye not many fathers: for in Christ Jesus I have begotten you through the gospel.")
    add_verse("[16] Wherefore I beseech you, be ye followers of me.")
    add_verse("[17] For this cause have I sent unto you Timotheus, who is my beloved son, and faithful in the Lord, who shall bring you into remembrance of my ways which be in Christ, as I teach every where in every church.")

    add_p("Stop livestreaming everything. You can't be talking to the whole world like you are talking to your local church.")
    add_p("There are certain things you don't say to the whole world.")
    add_p("In Rome, Paul goes straight to doctrine; he hardly went into personal details because Rome was different.")
    add_p("Not everyone you share the word of God with is your disciple.")
    add_p("Somebody liking what you are preaching, or listening to your audios, does not make the person a disciple.")
    add_p("Somebody saying, 'You are such a great blessing to me,' does not make them a disciple.")
    add_p("Paul had different relationships: Barnabas, Silas, Timothy, and the churches.")

    # Section 8
    add_h1("8. The Definition of a Disciple and Concluding Charge")
    add_p("Who is a disciple?")
    add_p("A disciple is a disciplined follower of a focused leader.")
    add_p("A disciplined follower of a focused leader.")
    add_p("Remember we started with sheep having no shepherd, and that sheep having no shepherd has a direction.")
    add_p("A disciplined follower of a focused leader.")
    add_p("This leader must be focused, number one, in his lifestyle.")
    add_p("In 1 Corinthians 4:17, Paul says, 'I have sent unto you Timotheus, who is my beloved son, and faithful in the Lord, who shall bring you into remembrance of my ways which be in Christ, as I teach every where in every church.'")
    add_p("It is not just doctrine; Timothy knows Paul's lifestyle.")
    add_p("Focused leadership will be seen in the conduct, number one.")
    add_p("Number two, in the message.")
    add_p("And number three, in your relationships.")
    add_p("When you choose to make disciples, it is a lifetime job.")
    add_p("You want to lead the person all life long. A lifetime job.")
    add_p("It doesn't mean the person will always be in your church. Some will be angry and walk away; some will even fight you.")
    add_p("Make sure you are still leading, and you are leading aright.")
    add_p("Who is a disciple? A disciplined follower of a focused leader.")
    add_p("Jesus taught everybody deep insights and revelations. In Matthew 7:28–29, they were astonished at his doctrine, for he taught them as one having authority, and not as the scribes.")
    add_p("That is not how it works for disciples. To disciples, he said, 'Follow me.'")
    add_p("Disciples don't just get information; they obey instructions.")
    add_p("That somebody sits in your meetings is not enough. That somebody hears you preach is not enough.")
    add_p("It is the people who follow your instructions.")
    add_p("After they had been to Jesus's meetings and liked his teaching, he said, 'Follow me,' and they dropped all they had and followed him.")
    add_p("If you are going to do that to people, then you must take responsibility.")
    add_p("You can't guess. You can't experiment. You must be focused.")

    add_ref("Matthew 7:28–29 (KJV)")
    add_verse("[28] And it came to pass, when Jesus had ended these sayings, the people were astonished at his doctrine:")
    add_verse("[29] For he taught them as one having authority, and not as the scribes.")

    add_h1("9. Closing Prayer and Consecration")
    add_p("Stand to your feet. Are you worth following?")
    add_p("Say: 'I am worth following.'")
    add_p("Say it loud: 'I am worth following.'")
    add_p("Shout and say: 'I am worth following!'")
    add_p("Pray for yourself. If he has called you to make disciples, he has called you to lead others with your life, with your message, with your relationships, and with your decisions.")
    add_p("Pray for yourself. Pray: 'Lord, make my life worth following.'")
    add_p("A disciple is a disciplined follower of a focused person, a focused teacher, a focused instructor, a focused shepherd.")
    add_p("Are you lacking focus? Are you lacking stability? It will be a problem for others.")
    add_p("Are you lacking stable conviction and direction in life? It is going to affect others.")
    add_p("Pray: 'My life is worth following. My life is worth engaging. My life is leading others well.'")
    add_p("Amen.")

def main():
    root_dir = os.path.dirname(os.path.abspath(__file__))
    raw_dir = os.path.join(root_dir, "Workers in Training Convention Transcript", "Raw")
    os.makedirs(raw_dir, exist_ok=True)

    # Document 1: Root Clean & Structured Docx
    root_docx_path = os.path.join(root_dir, "Workers in Training Convention 2026 – What to do with a new convert - Track 1 (Structured & Clean).docx")
    root_txt_path = os.path.join(root_dir, "Workers in Training Convention 2026 – What to do with a new convert - Track 1 (Structured & Clean).txt")

    # Document 2: Raw Folder Docx & Txt
    raw_docx_path = os.path.join(raw_dir, "Workers in Training Convention 2026 – What to do with a new convert - Track 1_Raw.docx")
    raw_txt_path = os.path.join(raw_dir, "Workers in Training Convention 2026 – What to do with a new convert - Track 1_Raw.txt")

    doc = create_document()
    txt_lines = []

    print("Building transcript content...")
    build_content(doc, txt_lines)

    # Save Root Docx
    print(f"Saving {root_docx_path}...")
    doc.save(root_docx_path)

    # Save Raw Docx
    print(f"Saving {raw_docx_path}...")
    doc.save(raw_docx_path)

    # Save Text Files
    full_text = "\n\n".join(txt_lines)
    with open(root_txt_path, "w", encoding="utf-8") as f:
        f.write(full_text)
    print(f"Saved {root_txt_path}")

    with open(raw_txt_path, "w", encoding="utf-8") as f:
        f.write(full_text)
    print(f"Saved {raw_txt_path}")

    print("All transcript deliverables generated successfully!")

if __name__ == "__main__":
    main()
