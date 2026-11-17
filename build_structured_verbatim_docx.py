# -*- coding: utf-8 -*-
"""
Builds the structured verbatim transcript for:
In the Spirit by the Spirit – Power Gifts - Track 2

Requirements:
- 100% verbatim spoken words from the audio (no rewrites, no summaries, no omissions)
- Clean, natural, well-spaced paragraphs (not cramped)
- Bible texts situated in their OWN separate paragraphs and formatted in BOLD
- Outputs both .docx and updated .txt
"""

import os
import docx
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH

# List of paragraphs: (text, is_bible_text)
PARAGRAPHS = [
    # Block 1
    ("Open your word this morning, that our minds receive wisdom, that everyone be wiser upon hearing the truth of your word this morning, in the name of Jesus. Amen.", False),
    ("Oh, praise the Lord! 2 Corinthians 1. In the Spirit, by the Spirit, this series started last year March, I think, and we've been looking at the things of the Spirit generally. Sometimes, spiritual things could take a long time, longer than you thought, but we're blessed anyway. Verse 18, 19 rather:", False),
    ("'For the Son of God, Jesus Christ, who was preached among you by us, even by me and Silvanus and Timotheus, was not yea and nay, but in him was yea. For all the promises of God in him are yea, that's what yea, and in him Amen, unto the glory of God by us. Now he which stablisheth us with you in Christ, and hath anointed us, is God; who hath also sealed us, and given the earnest of the Spirit in our hearts.'", True),
    ("Can you say amen? Romans 15:8, he says:", False),
    ("'Now I say that Jesus Christ was a minister of the circumcision for the truth of God, to confirm the promises that were made unto the fathers.'", True),
    ("Praise the Lord! And so we confirmed as what established. And so we said, quickly let's just move faster, that salvation, when you say salvation and you say the Spirit, right, you're saying the same thing.", False),
    ("There cannot be salvation without the giving of the Spirit. The giving of the Spirit is salvation. Are you listening?", False),
    ("Now, there is a study we're probably going to do in the second service that may help our understanding, because you will find the word 'giving of the Spirit' used in the Book of Acts that seems to differ from what I just said.", False),
    ("But anyway, salvation is synonymous with the giving of the Spirit. That is, you cannot be saved and not have the Spirit of God. Salvation is the giving, we can say the gift of the Spirit is salvation.", False),

    # Block 2
    ("Because now, I want you to pay attention to something, and you pay attention to everything. I said they are synonymous terms.", False),
    ("Now you know last week we looked about Jesus, we said Jesus is sacrifice, and we said most of the time our minds always stop at the cross. Remember I said that?", False),
    ("You know I said that, that we think Jesus died for my sins, and really the death of Jesus is one of the events of redemption. I'll take that again: the death of Jesus is one of the events of redemption.", False),
    ("His death, now listen carefully now, his death is sacrificial, but his death is not the sacrifice for salvation. I'll take that again: his death is sacrificial, in other words he suffered, but his death is not the sacrifice for salvation. Make sure that you listen to me well for this, so that you don't think I said what I didn't say, right?", False),
    ("Now, we said, and I said last week, that what Jesus did is evidenced by his Spirit that he gave us. Now let me go back to what I said earlier on, that his death is one of the events.", False),
    ("There are three events that mark the gospel of Christ, three: 1 Corinthians 15:3 and 4 states them. He said:", False),
    ("'how that Christ died for our sins according to the scriptures; and that he was buried, and that he rose again the third day according to the scriptures.'", True),
    ("Notice that he seems to, that Paul seems to, okay he died, okay, like one event on its own, then says he was buried and raised again the third day according to the scriptures. He didn't say he was buried according to the scriptures, then he was raised again according to the scriptures. Are you listening? Are you listening?", False),
    ("Because if you die, really you'll be buried anyway. That he was buried was proof that he died, okay? That he rose from the dead, okay, means that he was buried, you get that? You understand that?", False),
    ("Okay, so he has two or three events, right, that mark the gospel. So the gospel is not that Christ died; the gospel is Christ died, he was buried, and he rose from the dead. So are you listening? Are you listening? And that's how Paul puts it.", False),

    # Block 3
    ("Now, notice something in 1 Corinthians 15, go there quickly. 1 Corinthians 15 verse 14. Now he says, some folks are arguing whether Jesus rose from the dead or not, or sorry, how the dead rise, sorry, let me start from there. And he says:", False),
    ("'And if Christ be not what?' 14, 'be not what? Risen, okay, let's go on, our preaching is vain, and your faith is also vain.'", True),
    ("That word vain is fruitless, temporary. Now look at 17, let's take 17 together. 'And if Christ,' everybody let's read it together, let's go:", False),
    ("'be not raised, your faith is vain; ye are yet in your sins.'", True),
    ("Now he said, if Christ wasn't raised from the dead, your faith is vain, and you are still a sinner. He didn't say if Christ did not die. Are you listening? Are you listening?", False),
    ("So he says wasn't raised from the dead. So the rise or the raising or the resurrection of Jesus from the dead is the crown of redemption. It's the proof of redemption. It's the proof that man had been redeemed.", False),
    ("So you need to listen to that very well. If Christ wasn't raised from the dead, then you are yet in your sins. Praise the name of Jesus. Hallelujah.", False),

    # Block 4
    ("Okay, now let's quickly, we're going to move like this to come back this way, you must be ready to move that way. Now in John 3, look at the words of Jesus.", False),
    ("In John 3:3, verse 6, I knew that verse way before I got saved. Verily, verily, I say unto you, I used to wonder what verily verily meant, verily verily. Verily, verily, I say unto you, you there?", False),
    ("'Except a man be what? Except a man be born again, he cannot see the kingdom of God.'", True),
    ("Now notice, except a man be born again, he cannot see the kingdom of God. So Nicodemus asks in verse 4, what does it mean to be born again, that's basically what Nicodemus was saying. And look at what Jesus said:", False),
    ("'Verily, verily, I say unto thee, Except a man be born of water and of the Spirit, he cannot enter into the kingdom of God, right? So 6 says, That which is born of the Spirit is spirit, and that which is born of the flesh is flesh. Marvel not that I said unto thee, Ye must be born again. The wind bloweth where it listeth, and thou hearest the sound thereof, but canst not tell whence it cometh, and whither it goeth: so is every one that is born of the Spirit.'", True),
    ("So notice he uses the word 'born of the Spirit' how many times? Three times, right? So come on, how many times? So born of the Spirit explains born again. Is that clear? Is that clear?", False),
    ("But, come on, is that clear? Right. So that means, you know, he says you have to be born again. Now you don't give birth to yourself, and you don't determine that anyway, it's the parent that determines that anyway.", False),
    ("So he says be born of the Spirit. So Jesus calls it born of the Spirit.", False),

    # Block 5
    ("In John 4:14, he puts it differently. He says:", False),
    ("'whosoever drinketh of the water that I shall give him shall never thirst; but the water that I shall give him shall be in him a well of water springing up into eternal life,' or you call it everlasting life.", True),
    ("Then again John 7:38, 39. So Jesus in John 3 says born of the Spirit. Now in John 4:14, he says I'll give him, right, I'll give him, that is, I'll give to whoever believes, that's the word whoever drinks.", False),
    ("I'll give to him water, okay, that springs up unto everlasting life. I'll give him, okay, so if a man believes I'll give him, that is, he will have water like a well springing up unto everlasting life. So the first one says he's born, the second one says I'll give him.", False),
    ("Now John 7, John 7 in verse 37:", False),
    ("'If any man thirst, let him come unto me, and drink. He that believeth on me, as the scripture hath said, out of his belly shall flow rivers of living water. But this spake he of the Spirit, which they that believe on him should receive.'", True),
    ("So that means in salvation, listen, there is a new birth, right, and then something is received. You don't just get saved the way you are. That is, there is what is given. God gives something. And Jesus says he gives the Spirit, which is the same thing as you're born of the Spirit.", False),

    # Block 6
    ("Now go to John 14. John 14 and verse 16. He says:", False),
    ("'I will pray the Father, and he shall give you another Comforter, another Comforter, that he may abide with you forever; Even the Spirit of truth; whom the world cannot receive, because it seeth him not, neither knoweth him: but ye know him; for he dwelleth with you, and shall be where? In you. Shall be where? In you.'", True),
    ("Now I've told you many times, the language of the New Testament is what? IN. IN. Someone say IN.", False),
    ("So we said, you know we had two prophecies, remember that? Ezekiel's prophecy and Joel's prophecy. We said which one is closer to the New Testament? Ezekiel, because Ezekiel says IN. Joel says UPON. But Ezekiel says IN. IN is the language of the New Testament: in Christ, in you. Praise the Lord! Praise the Lord! Come on guys, are you here? Praise the Lord!", False),
    ("Remember, Paul never says to Philemon acknowledge every good thing that is upon you. Is that what he says? That is where? In you. In you in Christ. That's what you're supposed to find out.", False),
    ("So he says in you a well, you're born of the Spirit, he gave you his Spirit, right, because he will abide in you forever.", False),

    # Block 7
    ("Now we have two other statements: Luke's Gospel 24:48, 49. 49, Luke's Gospel 24:49, Luke says, Luke's Gospel 24:49, Luke says:", False),
    ("'you tarry in Jerusalem until you be endued with power from on high.' He says, 'I send the promise of my Father upon you.'", True),
    ("Then Acts 1:8, Acts 1:8, he says:", False),
    ("'you shall receive power after that the Holy Ghost is come upon you.'", True),
    ("And these are Jesus's words. And so that could get us a little bit confused, because we have in you, we have upon you. Is he talking about two different spirits or two different experiences of the same spirit?", False),
    ("Now I will advise you to stay close to the second service where I'm teaching on the power and revelation, and I'm going to explain a little bit of that in the language for and what it means.", False),
    ("Because sometimes you know in the study there's some traditional beliefs that have persisted for many years in the church world because the people who started them were very powerful leaders. And if someone is very powerful and he starts a doctrine, because many folks will hardly find out whether it's so or not, it can take a long time, a very long time.", False),
    ("Just like I told you the doctrine of rapture, the way it's said today. Somebody started it and then everybody just followed the train and nobody really asked the right questions. Praise the Lord! And quite a number of things like that. But you need to know why.", False),
    ("So when we say the Spirit upon you and the Spirit within you, and we can say well, the same Spirit but two experiences, but really not so. We'll find out that setting in the second service when we're looking at the power and revelation, because that's what sorts it out for us.", False),

    # Block 8
    ("Now, are you listening? Okay, good. Now so let's just take our mind a little bit away from how Luke describes that, because that may take us off the mark a little bit.", False),
    ("Now look at Paul in Romans 8. Paul is talking about the Spirit, right? In Romans 8, he says:", False),
    ("'The law of the Spirit of life where? In.' Notice Paul's language for the Spirit is IN. 'The law of the Spirit of life in,' someone say IN, 'in Christ Jesus.'", True),
    ("So the Spirit of life is where? In Christ Jesus. Now look at verse 8, are you there? Now again, this is not the thesis, but just pay attention. Again he mentions in, right, in verse 4, 'fulfilled in us.' Are you there? Are you there? Fulfilled in us.", False),
    ("So verse 8 says, 'those that are in the flesh, they that are in the flesh cannot please God.'", True),
    ("Now let's take 9 together. No, hold on again, don't forget: Romans 8:1:", False),
    ("'There is therefore now no condemnation to them which are where? In Christ.'", True),
    ("Now verse 2:", False),
    ("'The law of the Spirit of life what? In Christ.'", True),
    ("Very vital words: in Christ, in Christ. Verse 4, pardon me:", False),
    ("'fulfilled where? In us.'", True),
    ("So look at verse 9, let's go:", False),
    ("'But ye are not in the flesh, but where? In the Spirit.'", True),
    ("Now listen very well, study, you need to study it now: he says you are where? In Christ. He says the law of the Spirit of life where? In Christ. There is therefore now no condemnation to them who are where? In Christ. The righteousness of the law is fulfilled where? In us.", False),
    ("Listen: in Christ, in Christ Jesus, in us, then he says you are not where? In the flesh, but where? In the Spirit.", False),
    ("Now listen to that explanation: 'if so be that, uh, the Spirit of God what?' So the word in Christ and the indwelling of the Spirit, are they synonymous? Uh? Does the word in Christ mean the Spirit of God dwells in you? Come on, right? Come on, are you there?", False),
    ("'Now if any man have not who? The Spirit of Christ, he is none of his.'", True),
    ("Look at that: he's none of his. Now look at 9, 10 pardon me:", False),
    ("'And if Christ be where? Uh? In you.'", True),
    ("So that means in Christ and Christ in you are synonymous terms. If you are in Christ, it means Christ is in you. Are you there? Are you listening? Do you understand what I just did for you now? Okay, so when we say in Christ and Christ in you, it's the same truth that is being explained differently.", False),
    ("He says, 'And if Christ be in you, the body is dead because of sin; but the Spirit is life because of righteousness. But if the Spirit of him that raised up Jesus from the dead dwell where? In.'", True),
    ("Now look at how many times the word in you, in you, in is used again: Romans 8:1, Romans 8:2, Romans 8:4, Romans 8:9, Romans 8:10, Romans 8:11, where? In. Hallelujah! In you, in Christ.", False),

    # Block 9
    ("Now go further. He says:", False),
    ("'he that raised up Christ from the dead shall also quicken your mortal bodies by his Spirit that dwelleth where? In you.'", True),
    ("Look at 12:", False),
    ("'Therefore, brethren, we are debtors, not to the flesh, to live after the flesh,' that's a fact.", True),
    ("For if you,' now if you, conditional statement, right? 'If you, if you live after the flesh, you shall die.' You see, that statement, is it a possibility or an impossibility? I thought you know Bible interpretation, what is it? Uh? He says, 'brethren, can a man who is called a brother in Christ live after the flesh, eh? No.'", False),
    ("So the word if there is telling you, is painting a possibility that's an impossibility with the believer in Christ. Because he says to you, he says:", False),
    ("'if you live after the flesh, you shall die: but if you through the Spirit do mortify the deeds of the body, you shall live.'", True),
    ("I explained that word live to you before, right? Uh? Live different from walk, can you remember that? Good. So verse 14:", False),
    ("'For as many as are led,' in other words, verse 14 explains why you're not in verse 13, 'For as many as are led by the Spirit of God, they are the sons of God, okay?", True),
    ("Verse 15 tells you what led by the Spirit of God means. He says:", False),
    ("'For ye have not received the spirit of bondage again to fear; but ye have received the Spirit of adoption, whereby we cry, Abba, Father. The Spirit itself beareth witness where? With our spirit, that we are the children of God.'", True),
    ("Are you still there? Are you still there? Someone say IN. Okay, so that's the word IN.", False),
    ("So John's rendering, because you know Jesus spoke in different types of revelation to them, John's rendering is the closest to the New Testament, when he says he will dwell in you. That's the New Testament. He will dwell in you.", False),
    ("So Luke's explanation of upon is valid, but it's a part of the whole truth. Are you listening? Are you listening? Okay, let's move forward a bit.", False),

    # Block 10
    ("Now look at Galatians 4. Verse 4:", False),
    ("'But when the fulness of the time was come, God sent forth his Son, made of a woman, made under the law, to redeem them that were under the law, that we might receive the adoption of sons.'", True),
    ("Okay, let's take verse 6 together:", False),
    ("'And because ye are sons, God hath sent forth the Spirit of his Son into your hearts, crying, Abba, Father.'", True),
    ("The Spirit of his Son into your hearts, crying, Abba, Father. 1 Corinthians 2. Praise the Lord!", False),
    ("So salvation is to receive the Spirit. Is that clear? Is that clear?", False),
    ("Salvation is not just that you were transferred from darkness to light and then you can go back again, no. Salvation means you have received the Spirit of God. That's what it is, and you are in him.", False),

    # Block 11
    ("Now 1 Corinthians 2 verse 9:", False),
    ("'But as it is written, Eye hath not seen, nor ear heard, neither have entered into the heart of man, the things which God hath prepared for them that love him. But God hath revealed them unto us by his Spirit: for the Spirit searcheth all things, yea, the deep things of God. For what man knoweth the things of a man, save the spirit of man which is in him? even so the things of God knoweth no man, but the Spirit of God.'", True),
    ("I want us to read verse 12 together, let's go: 'Now we have received,' hold on, hold on. You know the first two verses were to explain something to you, okay? Now the verse 12 explains salvation. The first two verses do not explain salvation to you; the first two verses simply tell you, so God has revealed those things to us by his Spirit, the Spirit searcheth all things, the deep things of God, he says the things of God knows no man. He's talking about the Spirit there, he's not talking about salvation.", False),
    ("So in verse 12, verse 12 is why he explained verse 10 and 11. So let's take verse 12 again, let's go, come on: 'Now we have received,' hold on, he didn't say we will receive, uh? Come on, did he say we will receive? We have received: past tense, present tense, future tense? Uh? Which tense? So you have received. Say I have received! Past tense. Now, I have received what? Uh?", False),
    ("'Not the spirit of the world, but the spirit which is of God; that we might know the things that are freely given to us of God.'", True),
    ("Hallelujah! Hallelujah! In other words, your spirit knows. Your spirit already knows, that you might know. He didn't say that that we'll know, so you might know the things that are freely given to us of God. So salvation is to receive the Spirit of God. Hallelujah!", False),
    ("So what is salvation? It's to receive the Spirit of God. So you cannot be saved and not have the Spirit of God, because salvation is to receive the Spirit of God. Hallelujah! Are you listening to that?", False),
    ("So if you receive the Spirit of God, you don't, I had a theory that says when you're born again, okay, you are now in Christ, that's how they say it. But he said, okay let me put it the way he said it: he said you can get the dirt, you can take the man out of the dirt, so when you come out of the dirt, now I heard it from a very respectable preacher that I respected, so you take the man out of the dirt, that is salvation; so but it's time to take the dirt out of the man, that's sanctification and deliverance. You see what they say?", False),
    ("So you will need, they believe that you still have some very foreign spirits that lie somewhere in you, and so someone will feel you have a dual personality disorder. That is, you are in Christ, but not totally for Christ; that you are born again, but you have some other spirits that, you know, whether tenants or landlords, who walk together in your house, in your body, and then you, then sometimes your experiences don't help you.", False),

    # Block 12
    ("When I got saved, I used to fall sick. So I went to someone to counsel with me, so he asked me some questions. And he said I should give him some stories about my family, so I did give him some stories. Sometimes you help those kind of people to give you more unbelief.", False),
    ("So I told him that well, my family, my father told me this story, my uncle told me, he said, ah! He said, it is those spirits that are coming after you. Because some people believe that you can go to a water goddess and get children. And so they say, well, what were your parents doing before you were born? So well, they used to worship at the shrine. Ow! The guy shakes his head. You see, you're in trouble.", False),
    ("So what do you have to say? Well, they asked for you from the goddess. So the goddess gave you to your parents. So what's going to happen now? You have to go and beg them. Or you tell them to take their hands off you, something like that.", False),
    ("Some people believe that kind of stuff, and then they say go and ask your mother what happened when you were born. So that makes you home, you get near to mom, say mommy tell me the truth about how I was born. You know, and when you are an African, you know, it takes a lot of scriptural sense away from your mind, it's just really spooky.", False),
    ("You know, but in addressing this, I tell folks before, I said Abraham came from an idol-worshipping home. It's not likely his father prayed to God. You know that true? You know like, if your father was an idol worshipper, it's not likely to mention, he didn't even know who God, can't call him Jehovah, can't have called him El Shaddai, all those things. He just probably talked to the goddess or god he worshipped. And God didn't say, Abraham I wanted to talk to you, but you need deliverance.", False),
    ("The point is, everybody who's born into this world, from Paul's theology, the Bible says sin passed upon all men. So either way, you came from a goddess, you know what I mean anyway. I think it's just a wrong theology. A goddess can't give anybody a child, because they have no such capacity. Is that clear?", False),
    ("Some of you are happy because you heard that, because you thought that you came from, you came from under the rock. You know, so sometimes your experiences deceive you. And I know folks today that they are doubting their salvation because of certain habits they have.", False),
    ("Yes, you should walk in the word and break those habits, but you see, your salvation is not a product of your experiences. Your salvation is a gift from God. It's not a product of your experiences. Your experiences sometimes can be good, sometimes can be bad, sometimes they are just neutral. But salvation is a product of God's gift to you. Is that clear? Is that clear? Come on guys, is that clear? All right, good.", False),
    ("So he gave us his Spirit, so we are not, we're not, we're not, we are already, and I say it often, that it's the Spirit that God gave you that makes you holy. Some guys will say, well, the Holy Spirit can never come into a sinful, a sinful, a sinful vessel. So what does a sinful vessel do? Say the vessel has to become a righteous vessel first, a holy vessel, so the Holy Spirit comes here to do what? He's scared of what is not holy? No, no.", False),
    ("It's the Holy Spirit that comes into the life of a sinner and makes the sinner holy. God only gives his Spirit to sinners that believe. Amen! Is that clear? When you believe, he gives his Spirit to sinners. So that's the word of God. At salvation, you receive the Holy Spirit. Is that clear? Because that's what salvation is. Praise the name of Jesus. Is that clear?", False),

    # Block 13
    ("Now look at Ephesians 1. Ephesians 1:13. It says:", False),
    ("'In whom ye also trusted, after that ye heard the word of truth, the gospel of your salvation: in whom also after that ye believed,' notice this, 'ye were sealed with that holy Spirit of promise.'", True),
    ("Same thing in chapter 4 and verse 30:", False),
    ("'ye were sealed with the Holy Spirit of promise.'", True),
    ("That word sealed is used in the Roman Empire for that which is a proof of ownership, a proof of permanent ownership. So he says the Holy Spirit is a proof of permanent ownership. Are you there? Are you listening? It's a proof of permanent ownership.", False),
    ("So in salvation, God claims the sinner totally. He claims his past, he claims his present, and he claims his future. That's why it's the giving of the Spirit. You don't, you don't say, well, God didn't make us halfway, you came halfway, you came halfway. He made us through, through the way down into the future.", False),
    ("He has sealed us with his Spirit. Say I'm sealed with the Spirit! Now where is the Spirit? In you. So that's a seal, a proof of permanent ownership.", False),

    # Block 14
    ("Now let me ask you quickly. You know we said Jesus used three examples in John 3, John 4, John 7. He said receive. Now in John 4:14, he says everlasting what? Life. In John 3, he says born of the Spirit. In John 7, he says receive. How about John 16:12 and 13? He says, so he tells you it will dwell in you forever. I want you to listen to this. And we said being a seal is a proof of permanent ownership.", False),
    ("So Jesus says you are born of the Spirit, he says that which is born of the Spirit is what? Spirit. So in other words, listen to this: the Spirit of God is the DNA of the Christian. That's his heredity. That's why it's a permanent ownership. It's not, you didn't receive something that you will, you can give up.", False),
    ("It's like a birth. You can never, you can disown someone in the court because you know he got you angry, your son and all that, but you can never disclaim the fact that the child, the boy or the girl carries your life, or your DNA, or your genes. That can never be altered. Is that true? Is that true? Can never be altered.", False),
    ("So the man who has received the gospel carries in him and with him the very DNA of God, the Spirit of God. So the Spirit of God is not an experience. The Spirit of God is the life that we have from God. Are you listening? That's the life that we have from God. The way a child receives life from the parent, the life we receive from God is the Spirit of God.", False),
    ("Hallelujah! Let me see your hand if you understand that, let me see your hand, come on wave your hand. Is that clear to you? So it's not an experience; the Spirit of God is a life that we have of God.", False),

    # Block 15
    ("So John 16, however, you know in all of Jesus's teachings, we only find an intro into the truth, we don't find an explanation. In all of this. So you know we're discussing, don't forget, the gifts of the Spirit, and Jesus never mentioned the word gifts of the Spirit. He mentioned the word gift in John 4:10, so you saw that, but he doesn't mention gifts of the Spirit, just like Jesus never mentioned grace. Hallelujah! Hallelujah! Jesus never said, Jesus never said in Christ. Does he say in Christ? He didn't mention grace.", False),
    ("But look at John 16, he says there are many things that I desire to say to you, John 16:12:", False),
    ("'I have yet many things to say unto you, but ye cannot bear them now. Howbeit when he, the Spirit of truth, is come, he will guide you into all truth: for he shall not speak of himself; but whatsoever he shall hear, that shall he speak: and he will shew you things to come.'", True),
    ("So Jesus is telling you that in his teaching ministry, he gives an intro. You get that? He gives an intro, he doesn't give you the full picture.", False),
    ("So if you are discussing gifts of the Spirit, we're not going to get it in the four Gospels, but we can get an intro in the four Gospels. Praise the Lord! Are you there?", False),

    # Block 16
    ("Let's see an idea of where he begins to talk about the gifts of the Spirit. Again, salvation is what? Uh? You know you can call salvation forgiveness of sins, right? We can also call salvation what? Receiving the Holy Spirit, or the gift of the Spirit. Now don't forget, not gifts in plural, but just singular, right? The gift of the Spirit is salvation.", False),
    ("Now look at John 14 verse 12. John 14 verse 12. Are you there? Come on, are you there? What does he say? Uh? Verily, verily, that word comes up a lot. Verily, verily, I say unto you. Who is verily verily? Nobody called verily here? Verily, verily, I say unto you, let's go:", False),
    ("'He that believeth on me,'", True),
    ("hold on. Now he that believeth on him, what happens to he that believes? Uh? Uh-uh, please stop there. In the Book of John, he that believes, what happens to he that believes? He receives. Is that clear? Uh? Come on guys, is that clear? He that believes, in John 7:37 to 39, will receive the Spirit; in John 4:14, he that believes will receive the Spirit, right? Come on, isn't that right? So can we say, verily, verily, I say unto you, he that receives my Spirit, would that be correct? Is that correct? Okay, so go further: he that receives my Spirit, go further:", False),
    ("'the works that I do shall he do also, right? The works that I do shall he do also; and greater works than these; because I go unto my Father.'", True),
    ("And there's some very horrific interpretations of this verse. I know I've explained this many times to you, and the fellow said, what are the greater works? So what is greater? You hear what he says, he says, Jesus was saying greater works, he's trying to explain, he said Jesus worked miracles, he raised the dead, he healed the sick; Jesus is saying here, now this is how people they color their own imagination, he said, it's a theology, respected theology, he is a, he's a, he pastored the Metropolitan Tabernacle, the one started by Charles Spurgeon.", False),
    ("And so he said, now I'm saying this because I've written him already, so I'm not backbiting. I can't be backbiting on CD anyway, that is media-biting. So he said Jesus is saying here that the days of miracles will be over. And I found that interesting, because a greater work, what is greater than what Jesus did, that's what he said. Salvation of souls is a greater work.", False),
    ("And then someone who read the book, he's a pastor, he said I never saw it like that, and I said neither did I. Because I'm sure when you said, by God, revelation! I said if Jesus said a greater work, now I understand. He said greater works. How many salvations? Can you hear that? And I said oh, I never saw it like that, I said certainly, you always don't see anything.", False),
    ("And I said, and when I wrote him I said this, I said Jesus didn't say he that believes in me, the works I do he won't be able to do anymore. I said sir, even if I agree your greater work is salvation, he said the works I do he will first be doing, and greater works than this. So by the way sir, are we doing salvation? Amen. You decide who comes sir.", False),
    ("But the work there is ergon in the Greek, deeds, all right, a definite act, ergon, where you have energeo, words like that. So he says greater works than this, in other words he's pointing to a particular work: the works I do he will do also. Look at that word also, also! That means he'll raise the dead, he'll heal the sick, he'll cast out demons, and hallelujah, greater works than this, he didn't say a greater work, because I go unto my Father. Hallelujah! Glory to God! So you can see Jesus is speaking about the gifts. Now how do you know? Don't forget, the language of the New Testament is what? IN.", False),
    ("So read the next verse, let's take those verses together, next verse:", False),
    ("'And whatsoever ye shall ask in my name, that will I do, that the Father may be glorified where? In the Son. If ye shall ask any thing in my name, I will do it. If ye love me, keep my commandments.'", True),
    ("Don't forget, he didn't say keep the Ten Commandments. Is that clear? He didn't say keep the two commandments, new commandment. What's the commandments? He's giving them instructions, right? Keep my instructions. What instructions? What do our people say, talk about Jesus's commandments? What are they? Thou shall not worship any other god but me, is that what his commandments are? Verse 16:", False),
    ("'And I will pray the Father, and he shall give you what?' See, that's where he started from, can you see it? 'He shall give you another Comforter, that he may abide with you forever; Even the Spirit of truth; whom the world cannot receive, because it seeth him not, neither knoweth him: but ye know him; for he dwelleth with you, and shall be where? In you. I will not leave you comfortless: I will come to you. Yet a little while, and the world seeth me no more; but ye see me: because I live, ye shall live also. At that day ye shall know that I am where? In my Father, and ye in me, and I in you.'", True),
    ("So that's salvation. Can you see that? So in John 14, he didn't use the word gift, but he explains the working of the gifts, the workings of the supernatural. The believer can do it because he has received the Spirit, and he's going to work through him. Hallelujah!", False),

    # Block 17
    ("Now look at Mark 16. In Mark 16, he says:", False),
    ("'And these signs shall follow them that believe; In my name shall they cast out devils; they shall speak with new tongues; They shall take up serpents; and if they drink any deadly thing, it shall not hurt them; they shall lay hands on the sick, and they shall recover.'", True),
    ("Signs that follow, and he calls, Jesus calls them signs: cast out demons, speak in tongues, lay hands on the sick. So he calls them signs. Did he call them gifts? No, he didn't call them gifts. He calls them signs.", False),
    ("Now in the Book of Acts, we find the word gifts of the Spirit used for what? Now don't forget, in Acts 2:38, Peter said, 'Repent and be baptized every one of you in the name of Jesus Christ for the remission of sins, and ye shall receive the gift of the Holy Ghost.' What is the gift of the Holy Ghost there? Salvation. Acts 8:20, Simon the sorcerer offered money, and Peter said, 'Thy money perish with thee, because thou hast thought that the gift of God may be purchased with money.' That gift of God there is salvation, right?", False),
    ("So in the Book of Acts, the word gift of the Spirit is used for salvation. Don't forget that, the gift of the Spirit is used for salvation.", False),

    # Block 18
    ("Now, what did the Book of Acts call the manifestations of the Spirit? Are you listening? You need to know this, so that when you're reading Acts, you're not confused. In Acts 2:16, Peter said:", False),
    ("'this is that which was spoken by the prophet Joel; And it shall come to pass in the last days, saith God, I will pour out of my Spirit upon all flesh: and your sons and your daughters shall prophesy, and your young men shall see visions, and your old men shall dream dreams:'", True),
    ("Now notice what he used: pour out of my Spirit upon all flesh. Acts 8:15 and 16, Peter and John went to Samaria, 'prayed for them, that they might receive the Holy Ghost: (For as yet he was fallen upon none of them: only they were baptized in the name of the Lord Jesus.)' What word did he use? Fallen upon.", False),
    ("Acts 10:44, Cornelius's house, 'While Peter yet spake these words, the Holy Ghost fell on all them which heard the word. And they of the circumcision which believed were astonished, as many as came with Peter, because that on the Gentiles also was poured out the gift of the Holy Ghost. For they heard them speak with tongues, and magnify God.' Poured out, fell upon.", False),
    ("Acts 11:15, Peter explaining, 'the Holy Ghost fell on them, as on us at the beginning.' Verse 17, 'God gave them the like gift as he did unto us.' Acts 19:6, Paul at Ephesus, 'when Paul had laid his hands upon them, the Holy Ghost came on them; and they spake with tongues, and prophesied.' Came on them.", False),
    ("So the Book of Acts uses: poured out, fell upon, came upon, received the gift. For what? Speaking with tongues, prophesying, praise the Lord! That is, the utterance gifts! In the Book of Acts, the utterance gifts are described as the Spirit coming upon, falling upon, poured out. Are you listening? Are you listening? Good.", False),

    # Block 19
    ("Now what about revelation gifts? Word of knowledge, word of wisdom, discerning of spirits. How does the Book of Acts describe that? Does it say the Spirit fell upon them? No! Look at Acts 5:3, Peter said to Ananias, 'why hath Satan filled thine heart to lie to the Holy Ghost?' Verse 9, 'How is it that ye have agreed together to tempt the Spirit of the Lord?' What word is used? The Holy Ghost, the Spirit of the Lord.", False),
    ("Look at Acts 10. Peter is on the housetop praying, he falls into a trance, sees a sheet coming down with all manner of animals, four-footed beasts, creeping things, fowls of the air. And he was told, Peter, rise, kill and eat. And he says, no. Now did Peter say no unconsciously? No. You normally reflect your spiritual status in your dreams and visions. That's what you reflect.", False),
    ("Some of you have bad dreams a lot, and they come to pass. I know somebody that never has good dreams, never! I've never heard the person say you are blessed from the dream. It's either monkeys or dogs. When you start having dreams like that, you are living in fear. Because the Old Testament tells you that out of the multitude of business ariseth a dream. Are you listening? You wake up and you say, you're sweating, Jesus, Jesus, you know that's how you are.", False),
    ("Peter is saying no. Jesus took 40 days to explain the word to them, he comes back in a vision, Peter is still saying no. No, no, three times! I mean, even the psalmist says, once has the Lord spoken, twice have I heard, and three times the guy is saying no. Look at one thing: when he got up, the Bible says in verse 17, Peter doubted in himself what this, you know he's doubting in the vision. Hallelujah! I think I explained that to you, right? Lazarus, rich man, on Wednesday? Uh? Okay, let's move forward.", False),
    ("So in verse 19:", False),
    ("'While Peter thought on the vision, the Spirit said unto him, the Spirit said unto him, Behold, three men seek thee.'", True),
    ("So that the Spirit said to him, that's a revelation, right? Right? So come on, is that right? So two times. Now in Acts 10, sorry Acts 5:4 and 9, it's the Spirit. Now again, the Spirit said. Notice what's used for revelation is THE SPIRIT. Come on guys, are you following? Hey, are you following this? Are you following this?", False),
    ("Look at chapter 11: Agabus, Agabus, Agabus, Agabus. Agabus, the unconscious grandfather of the charismatic movement:", False),
    ("'And there stood up one of them named Agabus, and signified by the Spirit,' that is, he showed them by the Spirit, 'that there should be great dearth throughout all the world: which came to pass in the days of Claudius Caesar.'", True),
    ("Now how would he show them by the Spirit? Did he say, okay, Spirit becomes a projector machine, and you know, he must have done what, said it by speaking, right? Come on guys, uh? Hey, are you listening? By speaking. So he showed them by the Spirit. Now is that revelation? So are you seeing the word the Spirit, by the Spirit, the Spirit said, is used for what? Thank you.", False),
    ("Okay, go to Acts 13. Are you there?", False),
    ("'As they ministered to the Lord, and fasted, the Holy Ghost said, Separate me Barnabas and Saul for the work whereunto I have called them.'", True),
    ("Two things: future and past, revelation. The Holy Ghost said. Are you listening? So the Holy Ghost said, or the Spirit said, or they showed by the Spirit, refers to what? Revelation. And Peter doesn't, there's no word like gift for that one: the Holy Ghost said. But utterance is the Holy Ghost upon. Are you listening? Are you following it? All right.", False),

    # Block 20
    ("Look at Acts 20 verse 22. You know with these terms, I'm clarifying it for you so that when you're studying them, you understand it. Is that clear? Acts 20:22:", False),
    ("'And now, behold, I go bound in the spirit unto Jerusalem, not knowing the things that shall befall me there: Save that the Holy Ghost witnesseth in every city, saying that bonds and afflictions abide me.'", True),
    ("Now the Holy Ghost said, now the Holy Ghost said here will be what again? Uh? Revelation. Now go to 21, I hope you are paying attention. 21 verse 4:", False),
    ("'And finding disciples, we tarried there seven days: who said to Paul through the Spirit, that he should not go up to Jerusalem.'", True),
    ("This is a very clear text. Now look at verse 11, that's Agabus again:", False),
    ("'And when he was come unto us, he took Paul's girdle, and bound his own hands and feet, and said, Thus saith the Holy Ghost, So shall the Jews at Jerusalem bind the man that owneth this girdle, and shall deliver him into the hands of the Gentiles.'", True),
    ("So the Holy Ghost said is revelation in the Book of Acts, correct? But how do we receive that revelation? By what? Utterance. It appears that the church in Acts graduated from the utterance gifts to, not graduated, increased from the utterance gifts to the revelation gifts. So that means within the revelation gifts, all right, you will have the utterance gifts.", False),
    ("Now did you observe something? Okay, whilst Peter called the utterance gifts the gift, doron, there was no particular name or explanation for the, for Paul pardon me, for the revelation gifts. However, if you don't get this one, you should pray not just a four-line prayer, say Isaiah's prayers: that means a man who walks in the utterance gifts can equally walk in what? Revelation! Because it's about speaking. You get that?", False),
    ("Agabus spoke out revelation; the guys in Caesarea, they spoke out revelation; Agabus spoke out revelation; Acts 13:2, they spoke out by what? Revelation. All right, so the man who is into the utterance gifts can also speak revelation gifts. Praise the Lord! Is that clear? Is that clear? Okay.", False),

    # Block 21
    ("Now what about the power gifts? Miracles, signs, and wonders. You'll see that in the Book of Acts: miracles, signs, and wonders. So you may have the idea that there are different, maybe you call it, this guy has an anointing for, you know that's how they say it. They say, that man is, have you heard this, hear this very well: 'He host Reverend Somi with other anointed men of God.' Can you have a man of God who is not anointed? Well, in that context, yeah. Other anointed men of God. 'Ah that brother, he's anointed oh! When he's speaking like this, he speaks the very oracles of God. That sister, she sees things.'", False),
    ("There was one sister in our fellowship then, every fellowship time she has dreams. She's like Sister Joseph. Sister Joseph. You know Joseph had dreams, she had many many dreams, very very bad ones. And she'd be quiet: 'I sense that somebody here is in danger. I just sense you just have to pray.' That's why if you're a pastor, don't allow people to just come forward to give a word. If you're a pastor, you ask, what's the word?", False),
    ("You know where I learnt it from? There was this ministry headed by a friend of mine, now with the Lord now, we call him Chucky. So they used to have Holy Ghost meetings, so I used to go there, University of Ibadan, and yeah people give words, exhortation, prophecy, word of wisdom, word of knowledge. So they will write it down because pastor said they'll read it out.", False),
    ("So one day, a young sister, and I know this sister very very well, she's still my friend, so they will pass the paper across. Everybody was you know having a great time, laughing in the Spirit, having a great time, ministering to folks. So he just said, he wasn't paying attention, he just said, praise the Lord, yeah this is a word: 'All these things that you are doing, there are people speaking in tongues who are in hell right now! Keep your salvation.' I was watching and I know, okay, that's headless. Don't read things without seeing them first. Praise the Lord!", False),
    ("Sometimes some people get very offended. They say, pastor, that was the word for the church, take time, speak this for the church. They say, I will tell them. You tell me first. And if you want to say daddy, say it here, I am the pastor here. Pastor is the door of the sheep, all right? No, you don't go behind to enter a man's house, so let me be the one to allow you to enter or not to enter. That's why you press bells in houses. You don't just open the gate, 'Noah!' Say I'm here, can I come in? And you're already in. So you don't do that. Praise the Lord!", False),
    ("Even testimony, sometimes tell them what's the testimony. Let me tell you, church members get hurt. Somebody will give a testimony and say things like, 'Ah, let's be careful for the devil!' Oh my god! And they overcame the church by the word of the Satan's testimony! Praise the Lord! You learning something? All right.", False),

    # Block 22
    ("So that means the Book of Acts doesn't really clarify how those things happen. We can see they spoke it, we can see they heard things, we can see they, but we don't know how. Spirit upon. That's why it got many of us to think that you know I was in this meeting and the Spirit of God just came upon me, you see that?", False),
    ("Just came upon me, I didn't know what I was saying again. I was just seeing some things, I was just seeing some things. In fact, the kind of tongues I was speaking was strange, it was high tempo, like machine gun! I was just seeing some things that I didn't know that I was seeing these things, you know, I was just seeing some things. As I wanted to continue, the Spirit had gone! Hallelujah! It's because you read the Book of Acts.", False),
    ("1 Corinthians 12:1. And that's my teaching for today. Amen. For the last few weeks, I've just begun to have what I wanted to say, but you get blessed anyway, uh? All right.", False),

    # Block 23
    ("Now concerning spiritual gifts, Paul's speaking: things of the Spirit, brethren, I would not have you ignorant. So because that word there is pneumatikos, right, remember that? Pneuma, pneumatikos, that is, the things of the Spirit, what is given by the Spirit. So Paul is now to explain those things to us.", False),
    ("Now you shouldn't learn things primarily by experience; experiences will differ. As you see in the Book of Acts, things happened differently. But if you learn by doctrine, you'll be able to guide your experiences and explain them as well. So Paul calls them pneumatikos, and in verse 4, okay, so what Paul has done for us, Paul has taken the veil away from the Book of Acts so we can see what happened literally. That in verse 4, look at verse 4. Look at verse 4: 'Now,' come on let's take verse 4 together, let's go:", False),
    ("'there are diversities of gifts, but the same Spirit.'", True),
    ("What is the word gifts that Paul used there? Charisma, which is what? Charisma means an endowment, an ability, right? Charisma means an endowment, an ability. It's used for people, you know when you say that guy is gifted, right? He's gifted in speaking. When you see a football player like Lionel Messi, you know he is gifted. When you see Cristiano Ronaldo, he is gifted. When you see other players, they are players. No but I'm not joking.", False),
    ("You know there are some teams you play with and if Lionel Messi is in the team and they have just five players fit, they can still beat the other team. Just give him the ball often. You know if you're a footballer, if you've played football before, you take the ball from the defense, if you're like the distributor of the team, what you call the playmaker, you look around your team, who can I give the ball at this particular point in time? You see nobody. You look away. You see, I should stay in faith. So let me just, I should stay in faith, right?", False),
    ("But you see Lionel Messi or you see Cristiano Ronaldo, you just pass the ball to him, you don't care what's going to happen, because you know he's gifted! But there are some folks, you pass to them, you say move to the left, move to the right, because you're not even sure he knows what to do with it!", False),

    # Block 24
    ("So charisma is a natural endowment. It was used, it was used by the Greeks. In fact, Plato and Aristotle, who were major artists of words before Jesus came, they used it for natural talents. So charisma is what becomes natural. So we can say there are diversities of what is natural, but the same Spirit.", False),
    ("Don't forget, the Spirit of God has become the nature of the believer, can you remember that? So these are diversities. And he mentions three. You know, Howard Carter or Howard Carter majorly said to us that there are three groupings of the gifts of the Spirit, and I agree with him: they are the utterance gifts that say something, and then the revelation gifts that show something, and you have the power gifts that do something. Remember that? Remember that?", False),
    ("But the keyword, all right, the keyword is that these things are abilities. So the things we read in the Book of Acts were not experiences; they were abilities that produced experiences.", False),
    ("As you are right now, you are naturally endowed by the Spirit of God to do supernatural things, naturally! Now when I say naturally, I'm not referring to your human makeup, I'm referring to your spiritual makeup. Is that very clear? Come on guys, is that very clear? Good.", False),

    # Block 25
    ("Now let's just quickly talk about the gifts of healings quickly, because I hope I can rush through that. In verse 9, he says:", False),
    ("'To another faith by the same Spirit; to another the gifts of healing by the same Spirit.'", True),
    ("Now he uses the word gift, don't forget he's discussing gifts, he used the word gifts of healing three times: verse 9, that's the only gift he repeats the word gifts again, verse 9; and then verse 28, he said he set some in the church; then verse 30, 'have all the gifts of healing?' So it's a fact that the ability to heal, which is in Christ, is in diversities.", False),
    ("Now listen: so that means in the endowment for healing are diverse endowments. I'll explain to you what I mean. Now I like art. I like good dramas, very good ones, that don't necessarily have to have 'Babalawo' in them and all that, you know, but very good dramas that you know, and I really love to watch them. I remember one I watched recently, I loved the drama, October 1, quite good, because that guy, what's his name again, Kunle Afolayan, is quite one of the best produced in Africa, quite solid.", False),
    ("Now in watching drama, I have this aversion for a typecast. A typecast is somebody that if you see him as a bad guy, if he's doing good guy you think he's a bad good guy! Just a typecast, just one way. There are some folks, they can only act rich man.", False),
    ("When I used to do dramas, I almost became a typecast because I always acted the guy that will miss the rapture! Or a backsliding Christian, something like that, something terrible like that, backsliding Christian, miss the rapture, because they felt I could do it very well. You know, I could miss the rapture very easily! You want to pray for me? They used to pray for me.", False),
    ("Okay so, but there are some artistes, you know if you've seen what do you call it, Terminator, you have Arnold Schwarzenegger, Terminator 1 he was a bad guy, Terminator 2 he became the good guy. So he can act this way, he can act that way. That means that kind of actor is endowed to play different parts, though he's still acting, you get it?", False),
    ("Now the gifts of healing mean that the believer is endowed to minister healing differently. Sometimes you can get used to just one way. He's endowed to minister healing differently. Are you listening? Come on guys, are you listening? Okay.", False),

    # Block 26
    ("Now the first one we want to examine today, and probably a bit next week, is the laying on of hands. I've heard people teach how to receive healing, they say you receive healing by faith, by the word, there's laying on of hands and gifts of healing. It's ignorance, with due respect to those who teach that.", False),
    ("Now gifts of healing already shows you, you know we all said that it can't be gifts of healing based on the kind of sickness, can you remember that? Can you remember that? Good. We said it's gifts of healing because the word gifts already have to do with ministry, so it has to be the ministries to the sick in different ways. Do you understand that? Do you understand that? Okay.", False),
    ("So now, the first one we want to examine is laying on of hands. Laying on of hands is a gifts of healing. Mark 16:", False),
    ("'They shall lay hands on the sick, and the sick shall recover.'", True),
    ("They, that's the sick, I always like to say and the sick, because the problem there is the word they: they shall lay hands and they shall recover, you know, so I always like to say and the sick shall recover. I hope you understand that? Okay, thank you.", False),
    ("Now what does it mean to lay hands? Lay hands on the sick and the sick shall recover, what does it mean? Look at Matthew 8. Matthew 8. Here's Jesus. Matthew 8. In the Book of Matthew, the healing ministry of Jesus was mentioned, pay attention now, his healing ministry was mentioned in Matthew 4 verse 23 and 24, but how it was wasn't told.", False),
    ("He said they brought unto him verse 24 all sick people that were taken with divers diseases, torments, and those that were possessed, and they were lunatics, and the Bible says they were healed. He didn't tell us how. But in Matthew 8, we now find specifics.", False),
    ("He said and there was a man, okay, who was a leper, said 'Lord, if thou wilt, thou canst make me clean.' And Jesus put forth what? Uh? Verse 3, Matthew 8, he put forth what? And did what? Touched him. Notice that: he put forth his hand and touched, and said:", False),
    ("'I will; be thou clean.'", True),
    ("Now go somewhere further. In verse 14:", False),
    ("'And when Jesus was come into Peter's house, he saw his wife's mother laid, and sick of a fever.'", True),
    ("So the very first Pope had a wife! Is that not true? Very first Pope, he had a wife. 'And he did what? Uh?' Now pay attention: he touched what? Her hand.", False),
    ("'And the fever left her: and she arose, and ministered unto them.'", True),
    ("He touched her hand. You know the first one, we don't know what he touched, but the second time he touched her hand. Now many of us, I'm going to show you in a moment, many of us think laying on of hands is on the head. So even if someone says he has back pain, you put your hand on the head! Are you listening? Put your hand on the head. He said he touched her hand. Praise the Lord! Are you there?", False),
    ("He put his hand, just touched her hand, that's what he did. Now did he lay his hands on them? Did he lay his hands on her, uh, by what? Praise the Lord!", False),

    # Block 27
    ("Now go to 1 Timothy 5. 'Lay hands,' verse 22:", False),
    ("'Lay hands suddenly on no man.'", True),
    ("Now I've heard people say that word lay hands on no man means when you're in a meeting you can't just lay hands on people anyhow. Rubbish! Where it says lay hands suddenly on no man here, neither be a partaker of other men's sins, keep thyself pure: this was a verse they used for me when I laid hands on the sick in 1992, that I had polluted myself with people's sins, because I didn't know who they were.", False),
    ("It's good to know your Bible well enough. Paul's discussion here was about ordination. The word lay hands suddenly on no man is the word hastily, suddenly, hastily. It's the word tacheos in the Greek, t-a-c-h-e-o-s, it means to do it hastily. That is, you, for instance he said to him, do not ordain anyone to ministry except the person is blameless, husband of one wife, okay, is able to teach. So say lay hands suddenly on no man.", False),
    ("Look at Acts 13, Acts 13: Paul and Barnabas. Verse 3:", False),
    ("'And when they had fasted and prayed, and laid their hands on them, they sent them away.'", True),
    ("They laid their hands on them for ministry. Okay? Are you there? Are you following? Good.", False),
    ("So now, so the laying on of hands for ministry is not touch; it's one that involves to present people for ministry, which will involve you know what you call ordination, appointment, or consecration to the work of ministry.", False),

    # Block 28
    ("Now in laying hands on the sick, you'll discover that what Jesus did was to touch the sick. Mark 8:22. Are you listening? Are you listening? Okay, 22. Sorry, Mark 8, sorry for that, Mark 8:22:", False),
    ("'And he cometh to Bethsaida; and they bring a blind man unto him, and besought him to touch him.'", True),
    ("That's what, to touch him.", False),
    ("'And he took the blind man by the hand, and led him out of the town; and when he had spit on his eyes, and put his hands upon him, he asked him if he saw ought. And he looked up, and said, I see men as trees, walking.'", True),
    ("That means the man was getting healed gradually.", False),
    ("'After that he put his hands again upon his eyes, and made him look up: and he was restored, and saw every man clearly.'", True),
    ("Okay, so in touching the sick, which in more context we're going to see that, it's advisable, all right, to put your hands where, where? Where the pain is. Now I'm going to go to this later.", False),
    ("Now if you have a woman, and you're a man, and for instance you have some women that will tell you their pain is somewhere that is private, you don't say pastor said, uh-uh, no, no, brother, no! All right? There's a way to do it that I'm going to explain that to you.", False),
    ("But you see, he laid hands where there was the pain. He put his hands there, where there was the pain. So laying on of hands is not laying hands on the head, are you getting this? Are you following it? Now come back again: Mark 10.", False),

    # Block 29
    ("Mark 10. Verse 13:", False),
    ("'And they brought young children to him, that he should touch them: and his disciples rebuked those that brought them.'", True),
    ("Okay? And look at verse 16:", False),
    ("'And he took them up in his arms, put his hands upon them, and blessed them.'", True),
    ("So to touch. So that means it's a deliberate, so the word lay in the Greek is a deliberate touch.", False),
    ("Now you know there's a way, let me say someone so, Maiwada, where he's sometimes you want to leave the auditorium of the church, and he's moving before me, you are too efficient, can you calm down? And he's moving, and I'm like, excuse me. You know, I've touched him, but I'm not laying hands on him, but I touched him.", False),
    ("And I can also touch him like this: what is it now, you know, right? Am I not touching him? I can also touch him like this: go sit down! All right, so I can lay hands like that. You can you can also touch someone like this: how are you, and then you hug yourself, say how are you doing? So you can't say you're laying hands.", False),
    ("So the word lay hands is a deliberate touch. Now what is that deliberate touch for? It is to minister healing. It's a deliberate touch. You'll find out the word lay hands, you find that they laid hands upon Paul, upon Peter, the soldiers, they laid hands. It's also deliberate, you know I said deliberate touch. Not for turkey, they laid hands for what? To arrest! That's a deliberate touch.", False),
    ("So laying on of hands is a deliberate touch. While our own is to minister, we want to arrest the sickness! To minister to the sick. So it is not to put hands upon the head of people, do you understand it? Do you understand? It's not to put hands upon the head of people. It is a deliberate touch on the sick, which many times is on the part where there is the pain.", False),

    # Block 30
    ("Look at Mark 6. So laying on of hands is a deliberate touch. So if someone has a leg pain, where is it advisable I lay hands? On the leg. Praise the Lord! Amen. We're going to go into the specifics next week. But you see, it's a deliberate touch, that is, I am touching the person with something.", False),
    ("You see what I said? If Maiwada can be ill, praise God he's healed, but someone is coming to the service and I say how you doing, and we shake hands, and he goes. Then he comes back and says sir, I want you to pray for me, I'm sick in my body. I touched him earlier, and that was just how you doing. Now what do I do? Uh? I now deliberately, so that laying on of hands is to lay hands with the power of God. It's deliberate!", False),
    ("Now you know for instance, I don't think that when Peter's mother-in-law got up, she just looked at Jesus from afar. She must have you know moved close to him, served him food, and probably had said you're blessed.", False),
    ("And I don't think that when he touched her the second time she started shaking, ha brother master let me heal you, what's going on the power is here again, you know, no! Because that's not laying, even though he must have touched her, or shook her hand, but that's not, that's not a deliberate touch, do you understand that? Do you understand that? So laying on of hands is deliberate. Praise the Lord!", False),

    # Block 31
    ("Now see this: Mark 6. The effect, the effectiveness of laying on of hands can be seen here. In Jesus's own town, they actually dishonored him. The Bible says they were offended at him, in verse 3, Mark 6. And he said, Jesus said:", False),
    ("'A prophet is not without honour, but in his own country, and among his own kin, and in his own house. And he could there do no mighty work, save that he laid his hands upon a few sick folk, and healed them.'", True),
    ("So even in the midst of unbelief, laying on of hands can work! It still laid hands.", False),
    ("So in Matthew 20, sorry Acts 28 quickly, we're going to explain this a little bit more. Acts 28. This is Paul in captivity. Notice, he's in captivity, he's under imprisonment in a sort, because he's under arrest, he's being taken from place to place, yield to appear before Caesar. Now notice this: Acts 28 verse 7:", False),
    ("'In the same quarters were possessions of the chief man of the island, whose name was Publius; who received us, and lodged us three days courteously. And it came to pass, that the father of Publius lay sick of a fever and of a bloody flux: to whom Paul entered in, and prayed, and laid his hands on him, and healed him.'", True),
    ("Now notice something, pay attention: 'So when this was done,' what was done? Okay, good.", False),
    ("'Others also, which had diseases in the island, came, and were healed.'", True),
    ("Did you observe a similarity in the Matthew 8 account and this account? Like laying on of hands is the way to start. When Jesus laid hands in Matthew 8, verse 16 and 17, that's where many people now came and they were healed. The same thing here, I'm going to show you the sequence next week. So laying on of hands is a primary way to start ministering to the sick. I mean both of them: both Jesus and this man, Jesus was in a place he wasn't honored, Paul was in confinement, and he laid hands on the sick.", False),

    # Block 32
    ("So what is the endowment here? What is it? You know it says gifts of healing, right? So what's the endowment? Uh? Your hands? What do you call it? Ability to what? To put your hands on the sick? Everybody can do that! Doctors do it every time, giving injection, right? Say what's wrong with you, say what's wrong with you. Are they not laying hands? So what is the, what is the gift? To minister the power of God.", False),
    ("What exactly is the gift? Mantle? I didn't hear what he's saying. Uh? Coming through the body? I don't understand that. I got you, I was just kidding. Now it's a, laying on of hands will be what: a deliberate touch. Can you use your leg? And that is your hand. I trapped you! Praise the Lord! Okay, we'll get to that anyway.", False),
    ("So what is the gift there? Okay, go and think about it. I won't say it. I've said it already, I won't repeat it. So you can touch people deliberately, okay, to minister to the sick, to minister divine healing unto them.", False),

    # Block 33
    ("Now notice something, and we're close here now. Notice something, and I want to show you quickly: in Acts 3, in the case of Peter, notice how what they do is described. Acts 3, that man at the Beautiful Gate, Bible says:", False),
    ("'And he took him by the right hand, and lifted him up: and immediately his feet and ankle bones received strength.'", True),
    ("Where did the strength come from? It came through Peter's what? Uh? Through Peter's what? Touch, right? Now look at Acts 8 verse 5:", False),
    ("'Then Philip went down to the city of Samaria, and preached Christ unto them. And the people with one accord gave heed unto those things which Philip spake, hearing and seeing the miracles which he did.'", True),
    ("So that means the healing or the power is in your action. It's in your action! Lay hands on the sick. It didn't say go and declare the sick healed, say lay hands on the sick. Same thing like Paul: Paul healed them, that's what the Bible says, Acts 28. So it is in the action.", False),
    ("So that means the diversity is the action, diverse actions. Praise the Lord! Praise the Lord! Do you understand that? It's in the action.", False),

    # Block 34
    ("So in our studying the gifts of healing, we're going to study the action, we're not done with laying on of hands, we're going to continue next week. We're going to study the action and see how it's done. Notice in the instances, those folks received, they received from what? The touch. It's from the touch that they received. Praise the Lord!", False),
    ("You're blessed? Let's stand to our feet. Let's pray in the Holy Ghost, sing in the Holy Ghost, sorry. Sing in the Holy Ghost. [Singing in the Spirit / Tongues]", False)
]

def build_docx(output_path):
    doc = Document()
    
    # Configure page setup
    section = doc.sections[0]
    section.top_margin = Inches(1.0)
    section.bottom_margin = Inches(1.0)
    section.left_margin = Inches(1.0)
    section.right_margin = Inches(1.0)
    
    # Document Title
    title_p = doc.add_paragraph()
    title_p.paragraph_format.space_before = Pt(0)
    title_p.paragraph_format.space_after = Pt(4)
    title_run = title_p.add_run("In the Spirit by the Spirit – Power Gifts - Track 2")
    title_run.font.name = "Calibri"
    title_run.font.size = Pt(18)
    title_run.font.bold = True
    title_run.font.color.rgb = RGBColor(0, 51, 102)
    
    subtitle_p = doc.add_paragraph()
    subtitle_p.paragraph_format.space_before = Pt(0)
    subtitle_p.paragraph_format.space_after = Pt(18)
    subtitle_run = subtitle_p.add_run("Verbatim Spoken Audio Transcript (Well-Spaced & Structured)")
    subtitle_run.font.name = "Calibri"
    subtitle_run.font.size = Pt(11)
    subtitle_run.font.italic = True
    subtitle_run.font.color.rgb = RGBColor(110, 110, 110)
    
    # Process paragraphs
    for text, is_bible in PARAGRAPHS:
        p = doc.add_paragraph()
        p.paragraph_format.line_spacing = 1.15
        p.paragraph_format.space_before = Pt(4) if is_bible else Pt(0)
        p.paragraph_format.space_after = Pt(8)
        
        run = p.add_run(text)
        run.font.name = "Calibri"
        run.font.size = Pt(11)
        
        if is_bible:
            run.font.bold = True
            run.font.color.rgb = RGBColor(0, 0, 0)
        else:
            run.font.bold = False
            run.font.color.rgb = RGBColor(34, 34, 34)
            
    doc.save(output_path)
    print(f"Successfully created: {output_path}")

def build_txt(output_path):
    lines = []
    lines.append("IN THE SPIRIT BY THE SPIRIT – POWER GIFTS - TRACK 2")
    lines.append("Verbatim Spoken Audio Transcript (Well-Spaced & Structured)")
    lines.append("=" * 70)
    lines.append("")
    
    for text, is_bible in PARAGRAPHS:
        if is_bible:
            lines.append(f"**{text}**")
        else:
            lines.append(text)
        lines.append("")
        
    with open(output_path, 'w', encoding='utf-8') as f:
        f.write("\n".join(lines))
    print(f"Successfully created: {output_path}")

if __name__ == "__main__":
    docx_path1 = os.path.join("In the Spirit by the Spirit Transcript", "Raw", "In the Spirit by the Spirit – Power Gifts - Track 2_Raw.docx")
    docx_path2 = "In the Spirit by the Spirit – Power Gifts - Track 2_Raw.docx"
    txt_path = os.path.join("In the Spirit by the Spirit Transcript", "Raw", "In the Spirit by the Spirit – Power Gifts - Track 2_Raw.txt")
    
    build_docx(docx_path1)
    build_docx(docx_path2)
    build_txt(txt_path)
    print("All builds completed successfully!")
