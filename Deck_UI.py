import streamlit as st
from Main import main
#logo image
def streamlit_app():
    #title
    st.title('Risk Deck Automation')

    #header
    #st.header(" ")
    st.subheader("Please select the deck which you want to be automated")

    #radio button
    deck = st.radio("Select Deck: ",('None','Overall Deck','Customer Journey Deck','Fraud Analysis Deck'))

    if (deck == 'None'):
        st.info('No deck selected right now')

    elif (deck == 'Overall Deck'):
        st.success('Generating Overall Deck')
        main(ppt='Template Risk Deck - Overall.pptx')

    elif (deck == 'Customer Journey Deck'):
        st.success('Generating Customer Journey Deck')
        main(ppt='Template Risk Deck - customer journey.pptx')

    else:
        st.success('Generating Fraud Analysis Deck')
        main(ppt = 'Template Risk Deck - fraud analysis.pptx')

    #text
    st.markdown('You can access the Automated Deck from ***Output Folder***.')

if __name__ == '__main__':
    streamlit_app()
