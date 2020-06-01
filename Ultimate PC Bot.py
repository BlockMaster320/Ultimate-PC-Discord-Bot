#SETUP ooof
#Import Modules
import discord
from discord.ext import commands
from discord.ext import tasks
import os
from dotenv import load_dotenv
import math
import datetime
import _pickle
import xlwings

#Load the Token and Create a Bot
load_dotenv();
TOKEN = os.getenv('ULTIMATEPC_TOKEN');                                                                                  #load token from .env file
bot = commands.Bot(command_prefix = ';');

#Set Global Variables
authorizedList = ["BlockMaster", "TheGreatMurloc"];                                                                     #used for some command permissions \
guild = None;                                                                                                           #set global variables (than they are defined in on_ready event so we I don't have to getting them in every event repeatedly over and over again)
biTimeRange = 30;                                                                                                       #this variable represents from what time range of the image channel history will be the best images chosen
timeOffset = 2;                                                                                                         #just for displaying time information purposes (CET is shifted by 2 hours from UTC at the moment)

#Set Server's Specific Naming
dataFileName = 'Ultimate PC - DataFile';
dataFileBackupName = 'Ultimate PC - DataFile Backup';
excelFileName = 'Ultimate PC - Leaderboard';
dmHistoryFileName = 'Ultimate PC - DM History';

channelImagesName = 'u-wot-made';
channelGalleryName = 'gallery';
channelPcName = 'photo-challenges';
channelBiName = 'best-images';

emojiGeneralPointsName = 'PointGeneral'
emoji1pointName = 'MedalBronze';
emoji2pointName = 'MedalSilver';
emoji3pointName = 'MedalGold';
emojiMedalBronzeName = 'MedalBronze';
emojiMedalSilverName = 'MedalSilver';
emojiMedalGoldName = 'MedalGold';

roleCommand = 'Pretty High';
roleRank1Name = 'testRole1';
roleRank2Name = 'testRole2';
roleRank3Name = 'testRole3';
roleRank4Name = 'testRole4';

#Remove the 'help' Command To Be Set Later
bot.remove_command('help');


#GENERAL FUNCTIONALITY
#Tell When Bot is Connected to the Server and Set Global Variables
@bot.event
async def on_ready():
    #Set the Global Variables
    global guild;                                                                                                       #set global variable 'guild'
    guild = bot.guilds[0];

    global authorizedList;                                                                                              #set global list 'authorizedList'
    print(f"Members authorized to use some of the commands: {authorizedList}.");
    tempAuthorizedList = authorizedList.copy();
    authorizedList.clear();
    for personName in tempAuthorizedList:                                                                               #go trough the authorized list (containing just names) and convert it to a list of actual members (the member class)
        authorizedList.append(discord.utils.get(guild.members, name = personName));

    #Set Bot's Status
    await bot.change_presence(status = discord.Status.dnd, activity = discord.Game("MS Painting"));

    #Print Time Until the Next Best Images Update
    try:
        #Load Data from the DataFile
        dataFile = open(f'{dataFileName}.txt', 'rb');
        dataList = _pickle.load(dataFile);
        lastUpdateDateBi = dataList[1][0];

        #Print Time Info
        global biTimeRange;
        global timeOffset;
        currentTime = datetime.datetime.utcnow();
        print(lastUpdateDateBi);
        nextUpdateTime = lastUpdateDateBi + datetime.timedelta(days = biTimeRange + 7) + datetime.timedelta(hours = timeOffset);
        zeroBeforeMinute = "0" * (nextUpdateTime.minute <= 9);
        print(f"Time until the next BI update: {(nextUpdateTime - currentTime).days} d {(nextUpdateTime - currentTime).seconds // 3600} h {((nextUpdateTime - currentTime).seconds % 3600) // 60} min [{nextUpdateTime.date()}, {nextUpdateTime.hour}:{zeroBeforeMinute}{nextUpdateTime.minute}].");
    except:
        print("\033[91m" + "[DataFile doesn't exist (on_message)]" + "\033[0m");

    #Send a Message
    print(f"Bot {bot.user.name} connected to {guild.name}!\n___________________________________\n");

    """
    channelImages = discord.utils.get(guild.channels, name = 'main-chat');
    messageHistory = await channelImages.history(limit = None).flatten();
    print(len(messageHistory));
    """

#Send a Message to a Certain User
@bot.command(name = "sendMessage")
@commands.has_role(roleCommand)
async def send_message_to_user(ctx, user: discord.Member, *, message):
    #Send and Save the Message
    global guild;
    await user.create_dm();
    await user.dm_channel.send(message);

    messagePostDate = ctx.message.created_at + datetime.timedelta(hours = timeOffset);                                  #get time when was the message sent
    zeroBeforeMinute = "0" * (messagePostDate.minute <= 9);
    messageTimeInfo = f'{messagePostDate.date()}, {messagePostDate.hour}:{zeroBeforeMinute}{messagePostDate.minute}';
    print(f"[{messageTimeInfo}] to {user.name}: {message}");

    textFile = open(f'{dmHistoryFileName}.txt', 'a');                                                                   #save the message sent to the user to a .txt file
    textFile.write(f"[{messageTimeInfo}] to {user.name}: {message}\n");
    textFile.close();

#Save a Message Someone Sent to the Bot's DM
@bot.event
async def on_message(message):
    global guild;
    if (message.channel not in guild.channels) and (message.author != bot.user):
        messagePostDate = message.created_at + datetime.timedelta(hours = timeOffset);                                  #get time when was the message sent
        zeroBeforeMinute = "0" * (messagePostDate.minute <= 9);
        messageTimeInfo = f'{messagePostDate.date()}, {messagePostDate.hour}:{zeroBeforeMinute}{messagePostDate.minute}';
        print(f"[{messageTimeInfo}] from {message.author.name}: {message.content}");

        textFile = open(f'{dmHistoryFileName}.txt', 'a');                                                               #save the received message to a .txt file
        textFile.write(f"[{messageTimeInfo}] from {message.author.name}: {message.content}\n");
        textFile.close();
    await bot.process_commands(message);

#Delete Certain Amount of Messages
@bot.command(name = "deleteMessage")
@commands.has_role(roleCommand)
async def delete_messages(ctx, amount = 2):                                                                             #this will take as a argument amount of messages (if it's not specified it will be 2) and delete them
    await ctx.channel.purge(limit = amount);

#Compute "Simple" Math
@bot.command(name = "quickMath")
@commands.has_role(roleCommand)
async def quick_math(ctx, *, toCompute):
    await ctx.send(eval(toCompute));

#Send Bot's server-info Article Part
@bot.command(name = "sendText")
@commands.has_role(roleCommand)
async def send_mainInfo_text(ctx):
    await ctx.send("**Ultimate PC Bot.py**"
                   "\nHi! My major function here is to write your images data to the database needed to update the server's channels. Beside that I can afford you some pretty cool commands. (Prefix for the commands is ';'. You can find it next to the 'L' character on most of the keyboards.)"
                   "\n\n**;help**\n```This shows you a list of commands with a brief explanation what they do.```"
                   "\n**;info** @<user>\n```This function will show you an info embed with some data of the tagged user. (If no one is tagged, it will send your info embed.)```"
                   "\n;**images** (pc)<number>```Typing ';images' will show you list of all your registered images. Then you can get the url link of a specific image by typing ;images <number> (for any image) or ;images pc<number> (for a PC image).```"
                   "\nThe bot is still under construction and it's not running 24/7 at the moment. (You can access it at at least on Friday from 16:00 to 19:00 UTC). I'm not an experienced programmer so if you're interested and want to improve the bot the python code is available here");

#IMAGE MANIPULATION
#Handle Command Exceptions
@bot.event
async def on_command_error(ctx, error):
    errorCommand = ctx.command;

    if isinstance(error, commands.errors.CheckFailure):
        await ctx.send("You don't have permission to use this command. :( (Perhaps git gut?)");
    if isinstance(error, commands.errors.CommandNotFound):
        await ctx.send("Entered command doesn't exist. (But there are some which actually does!)");
    if isinstance(error, commands.errors.BadArgument):
        if (errorCommand == show_member_info):
            await ctx.send("You entered an invalid user as the argument. (Next time try to look for a person from *this* universe.)");
        if (errorCommand == show_member_images):
            await ctx.send("You entered an invalid image number as the argument. (Check your num lock, just in case. ;))");
    else:
        raise error;


@bot.command('help')
async def show_help(ctx):
    commandPrefix = bot.command_prefix;
    helpEmbed = discord.Embed(title = "Ultimate PC Bot.py | Commands", colour = discord.Colour.blue());
    helpEmbed.add_field(name = f'{commandPrefix}help', value = "Shows this embed.", inline = False);
    helpEmbed.add_field(name = f'{commandPrefix}info', value = "Shows your info embed.", inline = False);
    helpEmbed.add_field(name = f'{commandPrefix}images', value = "Gets links to your images.", inline = False);
    helpEmbed.set_thumbnail(url = bot.user.avatar_url);
    await ctx.send(embed = helpEmbed);

#Class Representing a Member
class EditMember():
    def __init__(self, memberID, memberName, memberTag):
        self.memberID = memberID;
        self.memberName = memberName;                                                                                   #this attribute will probably change eventually, it's just for case the member will leave the server and its name is needed (it's not possible to get a name of a member who isn't on the server (actually you can do it using member's message or discord.Member object but a message can be deleted and the discord.Member class cannot be pickled into the dataFile))
        self.memberTag = memberTag;

        self.imageNormalList = [];
        self.imagePcDict = {};
        self.pointsGeneral = 0;
        self.pointsPc = 0;

    def get_editMember_info(self):
        imageNormalNumber = len(self.imageNormalList);
        imagePcNumber = 0;
        for imageList in self.imagePcDict.values():
            imagePcNumber += len(imageList);
        infoList = [imageNormalNumber + imagePcNumber, imagePcNumber, self.pointsGeneral, self.pointsPc];
        return infoList;

    def print_editMember_data(self, memberName, memberID, imageNormalList, imagePcDict, pointsGeneral, pointsPc):
        print(f"memberName: {self.memberName}" * memberName,
              f"\nmemberID: {self.memberID}" * memberID,
              f"\nimageNormalList: {self.imageNormalList}" * imageNormalList,
              f"\nimagePcDict: {self.imagePcDict}" * imagePcDict,
              f"\npointsGeneral: {self.pointsGeneral}" * pointsGeneral,
              f"\npointsPc: {self.pointsPc}" * pointsPc);

#Prepare the DataFile
def create_dataFile():
    dataFile = open(f'{dataFileName}.txt', 'wb');

    lastUpdateDate = datetime.datetime(2019, 3, 29, 16, 57, 50, 0);
    lastUpdateDateBi = [datetime.datetime(2020, 5, 29, 17, 0, 0, 0), 0];
    memberDict = {};
    imageNormalList = [];
    imagePcDict = {};

    embedGalleryDict = {};
    embedPcDict = {};
    embedBiDict = {};

    dataList = [lastUpdateDate, lastUpdateDateBi, memberDict, imageNormalList, imagePcDict, embedGalleryDict, embedPcDict, embedBiDict];
    _pickle.dump(dataList, dataFile);

#Create an Embed for Embed Channel
def create_embed(message, attachment, messagePoints, emojiGeneralPoints, title = None):
    messageAuthor = message.author;
    messagePostDate = message.created_at;

    imageEmbed = discord.Embed(title = title, description = f"\n{messagePoints} \u200b {emojiGeneralPoints} \u200b \u200b | \u200b \u200b [jump to the message]({message.jump_url})\n\n{message.content}");
    imageEmbed.set_author(name = f"{messageAuthor.name}", icon_url = messageAuthor.avatar_url);
    imageEmbed.set_image(url = message.attachments[attachment].url);
    zeroBeforeMinute = '0' * (messagePostDate.minute <= 9);
    imageEmbed.set_footer(text = f"{messagePostDate.date()} | {messagePostDate.hour}:{zeroBeforeMinute}{messagePostDate.minute}  UTC");

    return imageEmbed;

#Get Points by Finding Specific Reactions
def get_generalPoints(message):
    messageReactions = message.reactions;
    messagePoints = 0;
    for reaction in messageReactions:
        if (type(reaction.emoji) == str):                                                                               #UNICODE emojis are strings so we need to don't let them enter the code which is working with discord.Emoji class (custom emojis on the server)
            continue;
        if (reaction.emoji.name == emoji1pointName):
            messagePoints += 1 * reaction.count;
        if (reaction.emoji.name == emoji2pointName):
            messagePoints += 2 * reaction.count;
        if (reaction.emoji.name == emoji3pointName):
            messagePoints += 3 * reaction.count;

    return messagePoints;

#Go Trough All the Images and Register Them to the DataFile
@bot.command(name = 'registerImages')
@commands.has_role(roleCommand)
async def register_images(ctx, createDataFile = False):
    #Get the Channels
    global guild;
    global biTimeRange;
    global timeOffset;
    global emojiGeneralPointsName;
    channelImages = discord.utils.get(guild.channels, name = channelImagesName);
    channelGallery = discord.utils.get(guild.channels, name = channelGalleryName);
    channelPc = discord.utils.get(guild.channels, name = channelPcName);
    emojiGeneralPoints = discord.utils.get(guild.emojis, name = emojiGeneralPointsName)

    #Load Data from the DataFile
    if (createDataFile):
        create_dataFile();

    dataFile = open(f'{dataFileName}.txt', 'rb');
    dataList = _pickle.load(dataFile);
    lastUpdateDate = dataList[0];
    lastUpdateDateBi = dataList[1];
    memberDict = dataList[2];
    imageNormalList = dataList[3];
    imagePcDict = dataList[4];
    embedGalleryDict = dataList[5];
    embedPcDict = dataList[6];
    embedBiDict = dataList[7];

    #Register the Images
    messageHistory = await channelImages.history(limit = None, after = lastUpdateDate, oldest_first = True).flatten();
    print(f"new message in the channel: {len(messageHistory)}");

    for message in messageHistory:
        if (message.attachments != []):
            #Get Message Information
            messageID = message.id;
            messageAuthor = message.author;
            messagePostDate = message.created_at + datetime.timedelta(hours = timeOffset);
            zeroBeforeMinute = "0" * (messagePostDate.minute <= 9);
            try:
                pcNumber = input(f"{messageAuthor.name} ({messagePostDate.hour}:{zeroBeforeMinute}{messagePostDate.minute}, {messagePostDate.day}/{messagePostDate.month}/{messagePostDate.year}): ");
            except:                                                                                                     #if there's any character that cannot be encoded in the member's name
                pcNumber = input(f"[name not readable] ({messagePostDate.hour}:{zeroBeforeMinute}{messagePostDate.minute}, {messagePostDate.day}/{messagePostDate.month}/{messagePostDate.year}): ");

            if (pcNumber == 'x'):                                                                                       #don't save the image if 'x' is entered as the pcNumber
                continue;
            elif (pcNumber != '') and (pcNumber[0] not in '123456789'):                                                 #if by a mistake something else than 'x' or a number is entered, save the image as a normal image
                pcNumber = '';

            #Save the Image
            if (messageAuthor.id not in memberDict.keys()):                                                             #search for the MemberEdit object or create a new one
                memberDict[messageAuthor.id] = EditMember(messageAuthor.id, messageAuthor.name.encode('utf-8'), messageAuthor.discriminator);
            editMember = memberDict[messageAuthor.id];

            for attachmentCount in range(len(message.attachments)):                                                     #in case that there are multiple attachments in the message
                if (pcNumber != ''):                                                                                    #add PC image to the MemberEdit object
                    pcNumber = int(pcNumber);
                    if (pcNumber not in editMember.imagePcDict.keys()):
                        editMember.imagePcDict[pcNumber] = [];
                    editMember.imagePcDict[pcNumber].append(messageID);
                    editMember.pointsPc += 1 - 0.5 * (len(editMember.imagePcDict[pcNumber]) > 1);

                    if (pcNumber not in imagePcDict.keys()):                                                            #add PC image to the general PC image dictionary
                        imagePcDict[pcNumber] = [];
                    imagePcDict[pcNumber].append(messageID);
                else:
                    editMember.imageNormalList.append(messageID);                                                       #add normal image to the MemberEdit object and general normal image list
                    imageNormalList.append(messageID);

            #Add Image's General Points According to Its Reactions
            messagePoints = get_generalPoints(message);
            editMember.pointsGeneral += messagePoints;

            #Send the Image to Channels
            for attachmentCount in range(len(message.attachments)):                                                     #in case that there are multiple attachments in the message
                galleryEmbed = create_embed(message, attachmentCount, messagePoints, emojiGeneralPoints);               #send the image to the Gallery Channel
                galleryEmbedMessage = await channelGallery.send(embed = galleryEmbed);
                embedGalleryDict[galleryEmbedMessage.id] = [messageID, messagePoints];

                if (pcNumber != ''):                                                                                    #send the image to the PhotoChallenges Channel
                    pcEmbed = create_embed(message, attachmentCount, messagePoints, emojiGeneralPoints, f"PhotoChallenge {pcNumber}");
                    pcEmbedMessage = await channelPc.send(embed = pcEmbed);
                    embedPcDict[pcEmbedMessage.id] = [messageID, messagePoints];

    #Save Data to the DataFile
    dataFile = open(f'{dataFileName}.txt', 'wb');
    lastUpdateDate = datetime.datetime.utcnow();
    dataList.clear();
    dataList = [lastUpdateDate, lastUpdateDateBi, memberDict, imageNormalList, imagePcDict, embedGalleryDict, embedPcDict, embedBiDict];
    _pickle.dump(dataList, dataFile);
    dataFile.close();

    #Make a Backup Save File
    dataFileBackup = open(f'{dataFileBackupName} [{lastUpdateDate.date()}].txt', 'bw');                                 #create a backup copy of the dataFile just in case its content gets lost
    _pickle.dump(dataList, dataFileBackup);
    dataFileBackup.close();

    print("\033[4m" + "\nImages have been registered." + "\033[0m");

#Post Last Month's Best Images from the Image Channel
@bot.command(name = 'updateBestImages')
@commands.has_role(roleCommand)
async def update_bestImages(ctx):
    #Get the Channels
    global guild;
    global biTimeRange;
    global emojiMedalBronzeName;
    global emojiMedalSilverName;
    global emojiMedalGoldName;
    channelImages = discord.utils.get(guild.channels, name = channelImagesName);
    channelBi = discord.utils.get(guild.channels, name = channelBiName);
    emojiGeneralPoints = discord.utils.get(guild.emojis, name = emojiGeneralPointsName);
    emojiMedalBronze = discord.utils.get(guild.emojis, name = emojiMedalBronzeName);
    emojiMedalSilver = discord.utils.get(guild.emojis, name = emojiMedalSilverName);
    emojiMedalGold = discord.utils.get(guild.emojis, name = emojiMedalGoldName);

    #Load Data from the DataFile
    dataFile = open(f'{dataFileName}.txt', 'rb');
    dataList = _pickle.load(dataFile);
    lastUpdateDateBi = dataList[1][0];
    biMonth = dataList[1][1];
    embedBiDict = dataList[7];

    #Create a List of Messages and Their General Points
    biMonth += 1;
    timeRange = datetime.timedelta(days = biTimeRange);
    messageHistory = await channelImages.history(limit = None, before = lastUpdateDateBi + timeRange, after = lastUpdateDateBi, oldest_first = True).flatten();
    messageList = [];
    messagePointsTotal = 0;
    for message in messageHistory:
        if (message.attachments != []):
            messagePoints = get_generalPoints(message);
            messagePointsTotal += messagePoints;
            messageList.append([message, messagePoints]);

    #Get the Best Images and Send Them to BestImages Channel                                                            #send an announcement
    biMonthEmbed = discord.Embed(title = f"BEST IMAGES OF THE MONTH {biMonth}", description = None, colour = discord.Colour.blue());
    biMonthEmbed.add_field(name = "Images Posted", value = f"{len(messageList)}");
    biMonthEmbed.add_field(name = "Total Points Earned", value = f"{messagePointsTotal}");
    biMonthEmbed.set_footer(text = f"\u200b \n from {lastUpdateDateBi.date()} to {(lastUpdateDateBi + timeRange).date()}");
    await channelBi.send(embed = biMonthEmbed);

    messageList.sort(key = lambda x: x[1], reverse = True);                                                             #sort the list by general points (second element of each sublist)
    lastMessagePoints = 0;
    messageCount = 1;
    for message in messageList:                                                                                         #get top 5 (or more) images from the list and send them to the channelBi
        if (messageCount <= 5) or (lastMessagePoints == message[1]):
            if (messageCount == 1):
                trophy = emojiMedalGold;
            elif (messageCount == 2):
                trophy = emojiMedalSilver;
            elif (messageCount == 3):
                trophy = emojiMedalBronze;
            else:
                trophy = "";
            biEmbed = create_embed(message[0], message[1], emojiGeneralPoints, f"#{messageCount} Image {trophy}");
            biEmbedMessage = await channelBi.send(embed = biEmbed);
            embedBiDict[biEmbedMessage.id] = [message[0].id, message[1]];

            lastMessagePoints = message[1];
            messageCount += 1;

    #Save Data to the DataFile
    dataFile = open(f'{dataFileName}.txt', 'wb');
    lastUpdateDateBi = lastUpdateDateBi + timeRange;
    dataList[1] = [lastUpdateDateBi, biMonth];
    dataList[7] = embedBiDict;
    _pickle.dump(dataList, dataFile);
    dataFile.close();

    print("\033[4m" + "\nBest images have been updated." + "\033[0m");

#Update Embeds in a Certain Channel
async def update_embedChannel(embedChannel, imageChannel, imageNumber, memberDict, embedDict, embedDictType, dataList, updateAuthor, emojiGeneralPoints):
    #Get Message Information
    messageHistoryGallery = await embedChannel.history(limit = imageNumber).flatten();
    for embedMessage in messageHistoryGallery:
        if (embedMessage.id in embedDict.keys()):                                                                       #the check is there just in case the embed message isn't in the embedDict (but it should be there (like it really should))
            imageMessageID = embedDict[embedMessage.id][0];                                                             #get message information from the EmbedDict
            imageMessageOldPoints = embedDict[embedMessage.id][1];
            try:
                imageMessage = await imageChannel.fetch_message(imageMessageID);                                        #get the Message from its ID
            except:                                                                                                     #continue if the message doesn't exist
                print("\033[91m" + "\n[message not found (update_embeds)]" + "\033[0m");
                continue;
        else:
            continue;

        #Update the Embed of the Message
        imageMessagePoints = get_generalPoints(imageMessage);
        embedEdit = embedMessage.embeds[0];
        if (imageMessagePoints != imageMessageOldPoints) or (updateAuthor):
            if (embedDictType == "gallery"):
                memberDict[imageMessage.author.id].pointsGeneral += imageMessagePoints - imageMessageOldPoints;         #add points to the memberEdit object (it has to happen only once, that's why there's the "gallery" check)
            embedDict[embedMessage.id][1] = imageMessagePoints;                                                         #update the points in the embedDict
                                                                                                                        #change points and update author info in the embed
            embedEdit.description = f"\n{imageMessagePoints} \u200b {emojiGeneralPoints} \u200b \u200b | \u200b \u200b [jump to the message]({imageMessage.jump_url})\n\n{imageMessage.content}";
            embedEdit.author.name = imageMessage.author.name;
            embedEdit.author.icon_url = imageMessage.author.avatar_url;
            await embedMessage.edit(embed = embedEdit);

    dataList[2] = memberDict;
    if (embedDictType == 'gallery'):
        dataList[5] = embedDict;
    if (embedDictType == 'pc'):
        dataList[6] = embedDict;
    if (embedDictType == 'bi'):
        dataList[7] = embedDict;
    return dataList;

#Update the Embeds Data (general points)
@bot.command(name = 'updateEmbeds')
@commands.has_role(roleCommand)
async def update_embeds(ctx, imageNumber = 200, updateAuthor = True):
    #Get the Channels
    global guild;
    channelImages = discord.utils.get(guild.channels, name = channelImagesName);
    channelGallery = discord.utils.get(guild.channels, name = channelGalleryName);
    channelPc = discord.utils.get(guild.channels, name = channelPcName);
    channelBi = discord.utils.get(guild.channels, name = channelBiName);
    emojiGeneralPoints = discord.utils.get(guild.emojis, name = emojiGeneralPointsName);

    #Load Data from the DataFile
    dataFile = open(f'{dataFileName}.txt', 'rb');
    dataList = _pickle.load(dataFile);
    memberDict = dataList[2];
    embedGalleryDict = dataList[5];
    embedPcDict = dataList[6];
    embedBiDict = dataList[7];

    #Update the Embeds
    dataList = await update_embedChannel(channelGallery, channelImages, imageNumber, memberDict, embedGalleryDict, "gallery", dataList, updateAuthor, emojiGeneralPoints);
    dataList = await update_embedChannel(channelPc, channelImages, imageNumber, memberDict, embedPcDict, "pc", dataList, updateAuthor, emojiGeneralPoints);
    dataList = await update_embedChannel(channelBi, channelImages, imageNumber, memberDict, embedBiDict, "bi", dataList, updateAuthor, emojiGeneralPoints);

    #Save Data to the DataFile
    dataFile = open(f'{dataFileName}.txt', 'wb');
    _pickle.dump(dataList, dataFile);
    dataFile.close();

    print("\033[4m" + "\nEmbeds have been updated." + "\033[0m");

#Update Rank Roles of All the EditMembers
@bot.command(name = 'updateRoles')
@commands.has_role(roleCommand)
async def update_roles(ctx):
    #Get the Rank Roles
    global guild;
    roleRank1 = discord.utils.get(guild.roles, name = roleRank1Name);
    roleRank2 = discord.utils.get(guild.roles, name = roleRank2Name);
    roleRank3 = discord.utils.get(guild.roles, name = roleRank3Name);
    roleRank4 = discord.utils.get(guild.roles, name = roleRank4Name);

    #Load Data from the DataFile
    dataFile = open(f'{dataFileName}.txt', 'br');
    dataList = _pickle.load(dataFile);
    memberDict = dataList[2];
    dataFile.close();

    #Update the Rank Roles
    for editMember in memberDict.values():
        pointsGeneral = editMember.pointsGeneral;
        member = discord.utils.get(guild.members, id = editMember.memberID);

        if (member is not None):                                                                                        #add a role according to EditMember's general points and delete all other roles
            if (pointsGeneral <= 1):
                await member.add_roles(roleRank1);
            if (pointsGeneral > 1) and (pointsGeneral <= 2):
                await member.add_roles(roleRank2);
                await member.remove_roles(roleRank1, roleRank3, roleRank4);
            if (pointsGeneral > 2) and (pointsGeneral <= 3):
                await member.add_roles(roleRank3);
                await member.remove_roles(roleRank1, roleRank2, roleRank4);
            if (pointsGeneral > 4):
                await member.add_roles(roleRank4);
                await member.remove_roles(roleRank1, roleRank2, roleRank3);

    print("\033[4m" + "\nRank roles has been updated." + "\033[0m");

#Update the Excel File Data
@bot.command(name = 'updateExcel')
@commands.has_role(roleCommand)
async def update_excel(ctx, lastPcNumber: int):
    #Load Data from the DataFile
    dataFile = open(f'{dataFileName}.txt', 'br');
    dataList = _pickle.load(dataFile);
    memberDict = dataList[2];

    #Load the Excel File
    excelFile = xlwings.Book(f'{excelFileName}.xlsx');
    excelSheet = excelFile.sheets[0];

    #Create Excel Table With Member Data
    global guild
    memberList = [];                                                                                                    #list containing EditMember's name, tag and pointsGeneral
    postedNormalImagesList = [];                                                                                        #list containing EditMember's total images posted
    pcList = [];                                                                                                        #list containing EditMember's PhotoChallenges
    memberCount = 0;
    for editMember in memberDict.values():                                                                              #write data to the lists
        memberCount += 1;
        member = discord.utils.get(guild.members, id = editMember.memberID);
        if (member is not None):                                                                                        #in case the member left the server (or was kicked or banned) we get its name from editMember object
            memberName = member.name;
            memberTag = member.discriminator
        else:
            memberName = editMember.memberName.decode('utf-8');
            memberTag = editMember.memberTag;

        memberList.append([memberCount, memberName, f"#{memberTag}", editMember.pointsGeneral]);
        postedNormalImagesList.append([len(editMember.imageNormalList)]);
        pcTempList = [None] * lastPcNumber;
        for pcNumber, pcImageList in editMember.imagePcDict.items():
            pcTempList[pcNumber - 1] = len(pcImageList);
        pcList.append(pcTempList);

    #Write the Data to the Excel File
    excelSheet.cells(2, 'A').value = memberList;
    excelSheet.cells(2, 'G').value = postedNormalImagesList;
    excelSheet.cells(2, 'J').value = pcList;

    print("\033[4m" + "\nThe Excel file has been updated." + "\033[0m");

#Move an EditMember's Image to a Different Location in the Object                                                       #USE ONLY IF THE EMBEDS ARE UPDATED! (otherwise it could be subtracting wrong values from the EditMember)
@bot.command(name = 'editImage')
@commands.has_role(roleCommand)
async def edit_image(ctx, messageID: int, action):
    #Get the Channels
    global guild;
    channelImages = discord.utils.get(guild.channels, name = channelImagesName);
    channelGallery = discord.utils.get(guild.channels, name = channelGalleryName);
    channelPc = discord.utils.get(guild.channels, name = channelPcName);
    emojiGeneralPoints = discord.utils.get(guild.emojis, name = emojiGeneralPointsName);

    #Load Data from the DataFile
    dataFile = open(f'{dataFileName}.txt', 'br');
    dataList = _pickle.load(dataFile);
    memberDict = dataList[2];
    embedGalleryDict = dataList[5];
    embedPcDict = dataList[6];

    #Get EditMember From the Message ID
    imageMessage = None;
    imageMessagePoints = None;
    try:                                                                                                                #get imageMessage and its points
        imageMessage = await channelImages.fetch_message(messageID);
        imageMessagePoints = get_generalPoints(imageMessage);
    except:
        print("\033[91m" + "\n[message not found (edit_message)]" + "\033[0m");
    editMember = memberDict[imageMessage.author.id];
    editMember.print_editMember_data(False, False, True, True, True, True);

    #Get the Image's Embeds
    embedGalleryMessage = None;
    embedPcMessage = None;
    for embedMessageID, imageMessageList in embedGalleryDict.items():                                                   #get gallery embedMessage
        if (messageID == imageMessageList[0]):
            try:
                embedGalleryMessage = await channelGallery.fetch_message(embedMessageID);
            except:
                print("\033[91m" + "\n[message not found (edit_message)]" + "\033[0m");
            break;

    for embedMessageID, imageMessageList in embedPcDict.items():                                                        #get PC embedMessage
        if (messageID == imageMessageList[0]):
            try:
                embedPcMessage = await channelPc.fetch_message(embedMessageID);
            except:
                print("\033[91m" + "\n[message not found (edit_message)]" + "\033[0m");
            break;

    #Remove the Image from the EditMember
    subtractGeneralPoints = False;                                                                                      #this variable will ensure that if there's no imageMessage with the ID it will not subtract general points from the EditMember
    if (messageID in editMember.imageNormalList):                                                                       #look for the image in EditMember's imageNormalList
        editMember.imageNormalList.remove(messageID);
        subtractGeneralPoints = True;
    else:
        pcKeyToDelete = None;
        for pcNumber, pcList in editMember.imagePcDict.items():                                                         #look for the image in EditMember's imagePcDict
            if (messageID in pcList):
                pcList.remove(messageID);
                subtractGeneralPoints = True;
                if (len(pcList) == 0):
                    editMember.pointsPc -= 1;
                    pcKeyToDelete = pcNumber;
                else:
                    editMember.pointsPc -= 0.5;
                break;
        if (pcKeyToDelete is not None):                                                                                 #delete imagePcDict's pcNumber key if there's no PC image remaining
            editMember.imagePcDict.pop(pcKeyToDelete);

    if (subtractGeneralPoints):
        messagePoints = get_generalPoints(imageMessage);
        editMember.pointsGeneral -= messagePoints;

    #Add the Image to a New Location in the EditMember
    if (action[0] in '123456789'):                                                                                      #add the image to EditMember's imagePcDict
        pcNumber = int(action);
        if (pcNumber not in editMember.imagePcDict.keys()):
            editMember.imagePcDict[pcNumber] = [];
        editMember.imagePcDict[pcNumber].append(messageID);

        messagePoints = get_generalPoints(imageMessage);                                                                #add general and PC points according to the image
        editMember.pointsGeneral += messagePoints;
        editMember.pointsPc += 1 - 0.5 * (len(editMember.imagePcDict[pcNumber]) > 1);

        if (embedPcMessage is not None):                                                                                #when moving a PC image to a different PC (just update the existing PC embed)
            embedEdit = embedPcMessage.embeds[0];
            embedEdit.title = f"PhotoChallenge {pcNumber}";
            await embedPcMessage.edit(embed = embedEdit);
        else:                                                                                                           #when moving a normal image to the imagePcDict (send a new PC embed)
            pcEmbed = create_embed(imageMessage, 0, imageMessagePoints, emojiGeneralPoints, f"PhotoChallenge {pcNumber}");
            embedPcMessageNew = await channelPc.send(embed = pcEmbed);
            embedPcDict[embedPcMessageNew.id] = [messageID, imageMessagePoints];
            if (embedGalleryMessage is None):                                                                           #when adding the image to a new EditMember object (send a new gallery embed and PC embed)
                galleryEmbed = create_embed(imageMessage, 0, imageMessagePoints, emojiGeneralPoints);
                embedGalleryMessageNew = await channelGallery.send(embed = galleryEmbed);
                embedGalleryDict[embedGalleryMessageNew.id] = [messageID, imageMessagePoints];

    elif (action == 'n'):                                                                                               #add the image to EditMember's imageNormalDict
        editMember.imageNormalList.append(messageID);
        messagePoints = get_generalPoints(imageMessage);                                                                #add general points according to the image
        editMember.pointsGeneral += messagePoints;

        if (embedPcMessage is not None):                                                                                #when moving a PC image to the imageNormalList (delete image's PC embed)
            embedPcDict.pop(embedPcMessage.id);
            await embedPcMessage.delete();
        elif (embedGalleryMessage is None):                                                                             #when adding a normal image to a new EditMember object (send a new gallery embed) (the elif is there just for the case we're trying to move a normal image to the imageNormalList for some reason - it has no effect then)
            galleryEmbed = create_embed(imageMessage, 0, imageMessagePoints, emojiGeneralPoints);
            embedGalleryMessageNew = await channelGallery.send(embed = galleryEmbed);
            embedGalleryDict[embedGalleryMessageNew.id] = [messageID, imageMessagePoints];

    elif (action == 'x'):                                                                                               #remove the image from the gallery
        if (embedGalleryMessage is not None):                                                                           #delete embed message id key (the embed was deleted)
            embedGalleryDict.pop(embedGalleryMessage.id);
            await embedGalleryMessage.delete();

        if (embedPcMessage is not None):
            embedPcDict.pop(embedPcMessage.id);
            await embedPcMessage.delete();

    #Save Data to the DataFile
    editMember.print_editMember_data(False, False, True, True, True, True);
    dataFile = open(f'{dataFileName}.txt', 'bw');
    dataList[2] = memberDict;
    dataList[5] = embedGalleryDict;
    dataList[6] = embedPcDict;
    _pickle.dump(dataList, dataFile);
    dataFile.close();

    print("\033[4m" + "\nThe image has been relocated." + "\033[0m");

#MEMBER COMMANDS
#Send an Info Embed of a Specific Member
@bot.command(name = 'info')
async def show_member_info(ctx, member: discord.Member = None, printData = False):
    #Load Data from the DataFile
    dataFile = open(f'{dataFileName}.txt', 'rb');
    dataList = _pickle.load(dataFile);
    memberDict = dataList[2];
    if (member is None):
        member = ctx.author;
    editMember = memberDict[member.id];

    #Get Data from a EditMember Function
    infoList = editMember.get_editMember_info();
    infoAllImageNumber = infoList[0];
    infoPcImageNumber = infoList[1];
    infoGeneralPoints = infoList[2];
    infoPcPoints = infoList[3];

    #Send a Member Info Embed
    if (not printData):
        infoEmbed = discord.Embed(title = f"{member.name} | Member Info", colour = discord.Colour.blue());
        infoEmbed.set_thumbnail(url = member.avatar_url);
        infoEmbed.add_field(name = "Total Images Sent", value = infoAllImageNumber, inline = True);
        infoEmbed.add_field(name = "PC Images Sent", value = infoPcImageNumber, inline = True);
        infoEmbed.add_field(name = "\u200b", value = "\u200b", inline=True);                                                #these blank lines has to be there to create 2x2 embed field grid (it's probably not possible any other way)
        infoEmbed.add_field(name = "General Points", value = infoGeneralPoints, inline = True);
        infoEmbed.add_field(name = "PC Points", value = infoPcPoints, inline = True);
        infoEmbed.add_field(name = "\u200b", value = "\u200b", inline=True);
        infoEmbed.set_footer(text = f"\u200b \njoined {member.joined_at.date()} | {member.display_name}");

        await ctx.send(embed=infoEmbed);

    #Print EditMember Data to the Console
    if (printData):
        editMember.print_editMember_data(False, False, True, True, True, True);

#Send an Embed Containing Links to all the Member's Images
@bot.command(name = 'images')
async def show_member_images(ctx, imageNumber = None):
    #Get the Channel
    global guild;
    channelImages = discord.utils.get(guild.channels, name = channelImagesName);

    #Load Data from the DataFile
    dataFile = open(f'{dataFileName}.txt', 'rb');
    dataList = _pickle.load(dataFile);
    dataFile.close();
    memberDict = dataList[2];
    editMember = memberDict[ctx.author.id];

    #Setup Strings For the Image Links and allImage List
    allImageFieldName = '';                                                                                             #setup strings containing the images links
    allImageString = '';
    pcImageString = '';
    sendEmbed = True;

    normalImageList = editMember.imageNormalList;                                                                       #setup a list containing all the images
    pcImageList = list(editMember.imagePcDict.values());
    allImageList = normalImageList;
    for pcList in pcImageList:                                                                                          #the list naming is little bit confusing here, pcImageList is imagePcDict' values converted to a list, these values are also lists (named pcList) containing the PC images themselves (pcImage)
        for pcImage in pcList:
            allImageList.append(pcImage);
    allImageList.sort();

    #Create String Containing Specific Images
    if (imageNumber is not None):
        #Create String Containing Specific PC Images
        if (imageNumber.startswith('pc')):
            pcNumber = int(imageNumber[2:]);
            if (pcNumber not in editMember.imagePcDict.keys()):                                                         #check if the EditMember has PC image with the pcNumber entered
                print("\033[91m" + "\n[message out of range (show_member_images)]" + "\033[0m");
                await ctx.send(f"You haven't posted any images for PhotoChallenge {pcNumber} yet.");
                sendEmbed = False;
            else:
                for pcImage in editMember.imagePcDict[pcNumber]:
                    try:                                                                                                #change message's embed description if the message doesn't exist
                        imageMessage = await channelImages.fetch_message(pcImage);
                        pcImageString += f"[PC {pcNumber}]({imageMessage.jump_url})\n"
                    except:
                        pcImageString += f"PC {pcNumber}\n"
                        print("\033[91m" + "\n[message not found (show_member_images)]" + "\033[0m");

        #Create String Containing a Specific Image
        elif (imageNumber[0] in '123456789'):
            allImageFieldName = "Image"
            if (int(imageNumber) > len(editMember.imageNormalList)):                                                    #check if the EditMember has that many images as the imageNumber entered
                print("\033[91m" + "\n[message out of range (show_member_images)]" + "\033[0m");
                sendEmbed = False;
                await ctx.send(f"You haven't posed that many images.");
            else:
                image = allImageList[int(imageNumber) - 1];
                try:
                    imageMessage = await channelImages.fetch_message(image);                                            #change message's embed description if the message doesn't exist
                    allImageString += f"[image {imageNumber}]({imageMessage.jump_url})"
                except:
                    allImageString += f"image {imageNumber}\n"
                    print("\033[91m" + "\n[message not found (show_member_images)]" + "\033[0m");

        else:                                                                                                           #if something else than a number or 'pc'number is entered raise an command error
            raise commands.errors.BadArgument;

    #Create String Containing All the Images and PC Images
    else:
        #Create String Containing All the Images
        allImageFieldName = "All Images";
        imageCount = 1;
        for image in allImageList:
            try:
                imageMessage = await channelImages.fetch_message(image);
                allImageString += f"image {imageCount}\n";
            except:
                allImageString += f"*[image {imageCount}]*\n"
                print("\033[91m" + "\n[message not found (show_member_images)]" + "\033[0m");
            imageCount += 1;

        #Create String Containing PC Images
        pcImageDict = editMember.imagePcDict;
        pcNumberList = sorted(pcImageDict.keys());                                                                      #due to the fact that a dictionary can't be easily sorted we firstly create a list of sorted keys (PC numbers) a then go trough the list and search for the images by the PC number in pcImagesDict
        for pcNumber in pcNumberList:
            for pcImage in pcImageDict[pcNumber]:
                try:
                    imageMessage = await channelImages.fetch_message(pcImage);
                    pcImageString += f"PC {pcNumber}\n";
                except:
                    pcImageString += f"*[PC {pcNumber}]*\n"
                    print("\033[91m" + "\n[message not found (show_member_images)]" + "\033[0m");

    #Send an Embed
    if (sendEmbed):
        imagesEmbed = discord.Embed(title = f"{ctx.author.name} | Images", colour = discord.Colour.blue());
        imagesEmbed.set_thumbnail(url = ctx.author.avatar_url);
        if (allImageString != ''):
            imagesEmbed.add_field(name = allImageFieldName, value = allImageString);
        if (pcImageString != ''):
            imagesEmbed.add_field(name = "PC Images", value = pcImageString);

        await ctx.send(embed = imagesEmbed);

#Run The Bot
bot.run(TOKEN);

