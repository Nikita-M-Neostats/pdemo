# Final (Without RAG) 
import azure.cognitiveservices.speech as speechsdk
import streamlit as st
import pandas as pd
import ast
from langchain import LLMChain
import regex as re
from langchain_core.prompts import PromptTemplate
from langchain_openai import AzureChatOpenAI
from docx import Document
from datetime import date, datetime
today = date.today()
time = datetime.now().time()
import warnings
warnings.filterwarnings("ignore")

region = st.secrets.region
api_key = st.secrets.api_key
llm = AzureChatOpenAI(
    api_version=st.secrets.api_version,
    model=st.secrets.model,
    temperature=0,
    api_key=st.secrets.api_key2,
    azure_endpoint=st.secrets.azure_endpoint,
)
# Set up speech configuration
speech_config = speechsdk.SpeechConfig(subscription=api_key, region=region)
speech_config.speech_recognition_language = "en-US"
speech_config.speech_synthesis_voice_name = "en-US-AvaNeural"
speech_synthesizer = speechsdk.SpeechSynthesizer(speech_config=speech_config)

# Use default microphone
audio_config = speechsdk.audio.AudioConfig(use_default_microphone=True)
speech_recognizer = speechsdk.SpeechRecognizer(speech_config=speech_config, audio_config=audio_config)

# Adjust the timeout properties
speech_recognizer.properties.set_property(speechsdk.PropertyId.SpeechServiceConnection_InitialSilenceTimeoutMs, "30000")  # 30 seconds
speech_recognizer.properties.set_property(speechsdk.PropertyId.SpeechServiceConnection_EndSilenceTimeoutMs, "2000")  # 2 seconds

def tts(text, voice, code):
    speech_config1 = speechsdk.SpeechConfig(subscription=api_key, region=region)
    speech_config1.speech_recognition_language = code
    speech_config1.speech_synthesis_voice_name = voice
    speech_synthesizer = speechsdk.SpeechSynthesizer(speech_config=speech_config1)

    result = speech_synthesizer.speak_text_async(text).get()
    if result.reason == speechsdk.ResultReason.SynthesizingAudioCompleted:
        print("Speech synthesized successfully.")
    elif result.reason == speechsdk.ResultReason.Canceled:
        cancellation_details = result.cancellation_details
        print(f"Speech synthesis canceled: {cancellation_details.reason}")
        if cancellation_details.reason == speechsdk.CancellationReason.Error:
            print(f"Error details: {cancellation_details.error_details}")


def recognize_speech(voice, code):
    speech_recognition_result = speech_recognizer.recognize_once_async().get()
    if speech_recognition_result.reason == speechsdk.ResultReason.RecognizedSpeech:
        recognized_text = speech_recognition_result.text
        with st.chat_message("user"):
            st.markdown(recognized_text)
            st.session_state.messages.append({"role": "user", "content": recognized_text})
        return recognized_text

    elif speech_recognition_result.reason == speechsdk.ResultReason.NoMatch:
        print(f"No speech could be recognized: {speech_recognition_result.no_match_details}")
        with st.chat_message("assistant", avatar='NeoStats_Logo_N.png'):
            tts("Sorry, I could not hear you.", voice, code)
            st.markdown("Sorry, I could not hear you.")
            st.session_state.messages.append({"role": "assistant", "content": "Sorry, I could not hear you."})
        return None 

    elif speech_recognition_result.reason == speechsdk.ResultReason.Canceled:
        cancellation_details = speech_recognition_result.cancellation_details
        print(f"Speech Recognition canceled: {cancellation_details.reason}")
        if cancellation_details.reason == speechsdk.CancellationReason.Error:
            print(f"Error details: {cancellation_details.error_details}")
        return None

def main():
    st.set_page_config(page_title="Neostats Life Insurance", page_icon="NeoStats_Logo_N.png")
    st.image('image.png', output_format="PNG", width=400)
    st.title("Welcome Call Agent")
    text = "An AI-driven virtual assistant designed to enhance customer interactions, providing personalized guidance to new customers."
    st.markdown(f"<h5>{text}</h5>", unsafe_allow_html=True)
    with st.sidebar:
        list_voice = ["en-US-AvaNeural", "en-US-AndrewMultilingualNeural", "en-IN-AaravNeural", "en-IN-AnanyaNeural"]
        voice = st.selectbox("Select a voice", list_voice)
        code=""
        if voice=="en-IN-AaravNeural" or voice == "en-IN-AnanyaNeural":
            code="en-IN"
        else:
            code="en-US"
        file=r"C:\Users\Nikita.Mate\POC\Prudential Insurance_Call\Prudential Insurance Company (Outbound call for newly acquire customers)\Insurance_records.xlsx"
        insurance_records = pd.read_excel(r"C:\Users\Nikita.Mate\POC\Prudential Insurance_Call\Prudential Insurance Company (Outbound call for newly acquire customers)\Insurance_records.xlsx")
        names = insurance_records['Name-Insurance'].tolist()
        name_insurance = st.selectbox("Select a customer name", names)
        st.session_state.insurance = insurance_records.loc[insurance_records['Name-Insurance'] == name_insurance, 'Insurance'].values[0]
        st.session_state.name = insurance_records.loc[insurance_records['Name-Insurance'] == name_insurance, 'Name'].values[0] 
        st.session_state.policy = insurance_records.loc[insurance_records['Name-Insurance'] == name_insurance, 'Policy'].values[0] 
        st.session_state.start_date = insurance_records.loc[insurance_records['Name-Insurance'] == name_insurance, 'Start Date'].values[0] 
        st.session_state.email = insurance_records.loc[insurance_records['Name-Insurance'] == name_insurance, 'Email ID'].values[0] 
        st.session_state.dob = insurance_records.loc[insurance_records['Name-Insurance'] == name_insurance, 'DOB'].values[0] 
        st.session_state.policy_no = insurance_records.loc[insurance_records['Name-Insurance'] == name_insurance, 'Policy number'].values[0] 
        
        st.session_state.policy_term = insurance_records.loc[insurance_records['Name-Insurance'] == name_insurance, 'Policy term'].values[0] 
        st.session_state.plan_name = insurance_records.loc[insurance_records['Name-Insurance'] == name_insurance, 'Plan Name'].values[0] 
        st.session_state.prem_mode = insurance_records.loc[insurance_records['Name-Insurance'] == name_insurance, 'Premium Mode'].values[0] 
        st.session_state.prem_amt = insurance_records.loc[insurance_records['Name-Insurance'] == name_insurance, 'Premium Amount'].values[0] 
        
        st.session_state.StartButton = st.button("Start Call")
        st.session_state.ActiveCall = False
        st.session_state.greeting_done = False
        if "messages" not in st.session_state:
                st.session_state.messages = []
        st.session_state.EndButton = st.button("End call")
    
        # Manage call state
        if st.session_state.StartButton and not st.session_state.get("ActiveCall", False):
            st.session_state.ActiveCall = True
            st.session_state.messages = []  # Clear messages for new session
            st.session_state.greeting_done = False  # Reset the greeting flag
            st.success("Call started.")
        elif st.session_state.EndButton and st.session_state.get("ActiveCall", True):
            st.session_state.ActiveCall = False
            st.session_state.greeting_done = False  # Reset the greeting flag
            st.success("Call stopped.")
            return

    if st.session_state.StartButton:
        voice_name=""
        if voice=="en-US-AvaNeural":
            voice_name="Jenny"
        elif voice=="en-US-AndrewMultilingualNeural":
            voice_name="Andrew"
        elif voice=="en-IN-AaravNeural":
            voice_name="Aarav"
        elif voice=="en-IN-AnanyaNeural":
            voice_name="Ananya"
        greeting = f"Good Afternoon! My name is {voice_name} and I’m calling from Neostats Insurance Company.  Am I speaking with {st.session_state.name}?"  
        if not st.session_state.get("greeting_done", False):
            tts(greeting, voice, code)
            st.session_state.messages.append({"role": "assistant", "content": greeting})
            st.session_state.greeting_done = True
            # Display the messages from session state
            for message in st.session_state.messages:
                with st.chat_message(message['role'], avatar='NeoStats_Logo_N.png'):
                    st.markdown(message['content']) 
        mesg = "Assistants Message : "+greeting
        prev_ans =None
        while st.session_state.EndButton == False: 
            recognized_text = recognize_speech(voice, code)
            template = """You are a call center assistant named Gennie for Neostats Insurance. Your role is to guide customers through the details of their insurance policy and answer thier question. You will be provided with list of Earlier chat messages between you (Assistant) and the customer (User), and recent messages.
                    If User's message is not related to policy matters, respond with: "Sorry, but I can only answer question related to your policy."
                    Before generating review the provided recent users_messages and recent assistants_messages. Check the assistant's last message and user's response to it. Based on that response, decide which option to choose and follow the guidelines below.
                    You must follow the steps outlined below each one at a time. You cannot cover all points in one go. After each step, check if the user has a clear understanding of that point and if they have any questions regarding it. If user has no question, it doesn't means that they are not interested; simply move to the next point. For Every step, go through entire chat messages and check if that point was covered in previous messages. If it was, move on to the next point.    
                    Step 1: If customer confirms that you are speaking with the intended person, thank them for confirmation and proceed. If the customer indicates that you have reached the wrong number or responds negatively, note it as wrong number, apologize for that, and politely end the conversation. 
                    
                    Step 2: Inform them that this is an important call regarding the policy that they recently obtained from us. Ask if it is a convenient time to speak with them.
                        1. If the customers indicates that they are busy and they cannot continue conversation then:
                            - Respectfully acknowledge their decision and ask them if you can schedule the callback. 
                            - If customer agrees to a callback, ask for preferred date and time. 
                            - If customer does not agree to a schedule callback, assure them that they can reach out anytime if they have questions or need assistance. 
                            - Close the conversation politely.
                        2. If the customer is willing to continue the conversation: 
                            - Welcome to Neostats and thank them for choosing {Insurance}. Ask them for their date of birth to verify the policy.
                            - If their date of birth matches perfectly (date month and year) with provided date, proceed. If date provided by customer is partial then ask them for full date of birth (date, month and year).
                            - If does not match, note the discrepancy and inform them that Date of Birth is not correct and end the call with a note apologizing for inconvinience. 
            
                    
                    Step 3: Confirm the customer's email ID. If correct, thank them for confirmation and proceed; If not, Inform them that you noticed that Email ID in our system is different from what they provided now and ask them if they want to update thier Email ID, If yes, then inform that they will shortly receive a link through SMS to provide their correct Email ID and proceed; If they do not want to update thier Email ID then simply proceed without saying thank you.
                        - Explain why an Email ID is essential in today's digital environment (e.g., for sending updates and documents).
                        - Ask customer to please confirm whether they have received the Insurance Policy documents sent to them. If customer's response is positive (e.g., "yes"), indicating that they have recevied the documents, ask for date of when they received documents. If the customer did not receive the documents, inform them that you have noted it down and will check the status of their documents. 
                    
                        - Inform the customer that a soft copy of the policy documents has been sent to their registered Email ID, and a link to view policy documents online has been sent via SMS to thier registered mobile number.
                    
                    Step 4: Guide the customer through the details of thier policy (Policy Number, Policy for, Policy Term, Plan Name, Premium Mode, Premium Amount, Start Date etc.) along with all this points tell few coverage inclusion and exclusion briefly.
                        - Ask if they have any question or concerns about these details.
                            - If the customer has any question or query, answer them.
                            - If the customer has any concern about the details, apologize for the inconvenience and assure them that thier concern is very important and will be addressed properly. Note thier concern and inform them that a representative will look into the details and get back to them. Ask for a callback date and time.  
                        - Ask if there is anything else you can assist them with.
                        - If no further questions:
                            - Ensure the customer feels supported and knows that help is always available.
                            - Thank them for their time and welcome them to the company.
                            - Close the conversation politely.
                    
                    Step 5: If the customer expresses disinterest in continuing the chat after some interaction:
                        - If the customer says "no", check what the recent assistant's message was.If it was related to continuing the conversation, end the call. 
                        
                    If the user's message is None, simply ask, "Are you there?"
                    
                    Response Formatting"
                    - Your response should be formatted as a Python dictionary with two keys:
                    1. "Response" whose value will be Your response to the customer.
                    2. "End_Call" whose value will be "yes" if the call is ending, otherwise "no".
                    
                    - Make sure not to give any backticks or other string expect dictionary which will starts and ends with curly brackets. 
                    - Avoid giving any leading or trailing whitespaces, unnecessary blank lines, or any other text or symbols besides the dictionary.4
                    - Avoid using customers name at from step 3.

                    Insurance Taken: {Insurance}
                    Date of birth: {DOB}
                    Email ID: {email}
                    Policy for: {policy}
                    Policy Number: {policy_no}
                    Policy Term: {policy_term}
                    Plan Name: {plan_name}
                    Premium Mode: {prem_mode}
                    Premium Amount: {prem_amt}
                    Start Date: {startdate}
                    Time: {time}
                    recent assistants_messages: {assist_msg}
                    recent users_messages: {users_msg}
                    chat history: {history}
                    """
           
            prompt = PromptTemplate(
                        template=template,
                        input_variables=["Insurance", "DOB", "email", "policy", "policy_no", "policy_term", "prem_mode", "prem_amt", "plan_name", "startdate", "time", "assist_msg", "users_msg", "chat_messages", "history"]
                        )
            chain = LLMChain(prompt=prompt, llm=llm)
            
            answer = chain.run({
                'Insurance': st.session_state.insurance,
                'policy': st.session_state.policy,
                'policy_no':st.session_state.policy_no, 
                'email':st.session_state.email,
                'DOB': st.session_state.dob,
                'policy_term': st.session_state.policy_term,
                'plan_name': st.session_state.plan_name,
                'prem_mode': st.session_state.prem_mode,
                'prem_amt': st.session_state.prem_amt,
                'startdate': st.session_state.start_date,
                'assist_msg': mesg,
                'time': time,
                'users_msg': recognized_text,
                'history': st.session_state.messages
            })
            ans = answer.replace("```python","").replace("```","")
            with st.chat_message("assistant", avatar='NeoStats_Logo_N.png'):
                Res_dict = ast.literal_eval(ans)
                ans = list(Res_dict.keys())[0]
                ans_value = Res_dict[ans]
                end_call = list(Res_dict.keys())[1]
                end_callvalue = Res_dict[end_call]

                tts(ans_value, voice, code)
                st.markdown(ans_value)
                st.session_state.messages.append({"role": "assistant", "content": ans_value})
                if ans_value == "Are you there?" and recognized_text == None and prev_ans != "Are you there?":
                    mesg = prev_ans
                elif ans_value == "Are you there?" and recognized_text == None and prev_ans == "Are you there?":
                    st.session_state.ActiveCall = False
                    st.session_state.greeting_done = False  # Reset the greeting flag
                    st.success("Call ended automatically.")
                    break
                else:
                    mesg = "Assistant message: " + ans_value
                prev_ans = ans_value
                if end_callvalue.lower() == "yes":
                    st.session_state.ActiveCall = False
                    st.session_state.greeting_done = False  
                    st.success("Call ended automatically.")
                    break 
    if st.session_state.EndButton==True:
        template1 = """
            You are given call transcript between an assistant and a customer, where the assistant is making a welcome call from an Insurance company to newly acquired customers. Your task is to analyze the transcript and provide information for the following attributes: 
            
            1. Determine if the correct person was contacted.
            2. Schedule_call: Check if the callback was scheduled.
            3. Schedule_time: Identify if a time and date were mentioned for the scheduled callback. Based on today’s date, provide the specific scheduled date (e.g., if the customer says "tomorrow" and today’s date is September 10th, 2024, then the scheduled date will be September 11th, 2024).
            4. Verify if the date of birth was validated.
            5. Confirm if the email ID provided was correct.
            6. Check if the customer received the documents.
            7. Determine if there were any concerns about the policy details. If so, summarize the concern.

            Response Formatting:
            - Your response should be formatted as a Python dictionary with five keys:
            1. "Schedule_callback" whose value should be "yes" if the customer was busy and have schedule the callback; "no" if the customer was busy and did not schedule the callback, and "Attended" if the customer had a discussion with the assistant.
            2. "Scdedule_time" whose value should be Scheduled time and date if callback was scheduled, otherwise "null". Date should be DD-MM-YY TIME format.
            3. "Document_receving_date" whose value should be Date when customer recevied documents  if mentioned, otherwise "null". Date should be DD-MM-YY format.
            4. "Comment" whose value should reflect any complaint or comment based on points below:
                - "Customer is Busy" if customer is busy and cannot continue conversation.
                - "Incorrect Email ID" if the email ID provided by customer does not match with provide one. 
                - "Wrong Number" if the wrong person was contacted or wrong number.
                - "Invalid Date of Birth" if the date of birth said by customer does not match to provided date.
                - "Documents not received" if the documents were not received by customer.
                - "None" if none of above is applicable.
            4. "Concern Summary": The value should be "None" if the customer did not express any concerns. If there were concerns, the value should be a summary of those concerns.
            5. "Email_ID Updation": whose value should be "Yes" if customer is willing to update thier email ID if it was wrong. otherwise "No"

            - Make sure not to give any backticks or other string expect dictionary which will starts and ends with curly brackets. 
            - Do not include any leading or trailing whitespaces, unnecessary blank lines, or any other text or symbols besides the dictionary.

            Call transcript: {call_transcribe}
            Today's date: {today}
            """
        prompt1 = PromptTemplate(
                    template=template1,
                    input_variables=["call_transcribe","today"]
                    )
        chain1 = LLMChain(prompt=prompt1, llm=llm)
        print("Call :", st.session_state.messages)
        answer1 = chain1.run({
            'call_transcribe': st.session_state.messages,
            'today': today
        })
        ans1 = answer1.replace("```python","").replace("```","")
        if st.session_state.messages:
            print(ans1)
            Res_dict = ast.literal_eval(ans1)
            reschedule_call = list(Res_dict.keys())[0]
            reschedule_call_value = Res_dict[reschedule_call]
            reschedule_time = list(Res_dict.keys())[1]
            reschedule_time_value = Res_dict[reschedule_time]
            doc_date = list(Res_dict.keys())[2]
            doc_date_value = Res_dict[doc_date]
            comment = list(Res_dict.keys())[3]
            comment_value = Res_dict[comment]
            concern_summary = list(Res_dict.keys())[4]
            concern_summary_value = Res_dict[concern_summary]
            emailid_updt = list(Res_dict.keys())[5]
            emailid_updt_value = Res_dict[emailid_updt]
            insurance_records.loc[insurance_records['Name-Insurance'] == name_insurance, 'Reschedule Call'] = reschedule_call_value
            insurance_records.loc[insurance_records['Name-Insurance'] == name_insurance, 'Reschedule Time'] = reschedule_time_value
            insurance_records.loc[insurance_records['Name-Insurance'] == name_insurance, 'Document Received Date'] = doc_date_value
            insurance_records.loc[insurance_records['Name-Insurance'] == name_insurance, 'Comment'] = comment_value
            insurance_records.loc[insurance_records['Name-Insurance'] == name_insurance, 'Concern Summary'] = concern_summary_value
            insurance_records.loc[insurance_records['Name-Insurance'] == name_insurance, 'Update Email ID'] = emailid_updt_value
            insurance_records.to_excel(file, index=False)

if __name__ == "__main__":
    main()
