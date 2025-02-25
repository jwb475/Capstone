# Evolving Dialogue: The Role of Gender in Analyst Interactions in Earnings Conference Calls
This project is a NYU Abu Dhabi Economics Capstone Project studying the impact of sentiments and gender in earning conference calls. It aims to understand if there is a correlation between different sentiments (tone, vaguneness, and modality), gender, and the market returns of a firm. The project elaborates builds upon "Sentiment Analysis and Gender Differences in Earning Calls" (Amicci 2018). The primary focus of our project is to enhance the understanding of interactions during the question and answer (Q&A) sessions of earnings calls by tagging speaker interactions. This enables the creation of interaction terms that can be further analyzed to explore how these dynamics affect sentiment and market reactions. 

# Methodology 
1. Data Collection: The transcripts of earnings conference calls were obtained from Capital IQ covering a period from 2004 to 2018. The data includes nearly 78,000 calls across 4,978 firms. The transcripts were separated into 2 sections: Presentation and Question & Answers. 
2. Tagging Interactions: Each speaker's contributions during the Q&A sessions were tagged to identify who is speaking to whom. This tagging allows for the construction of interaction terms that represent the dialogue between executives and financial analysts.
3. Sentiment Analysis: Using the Loughran and McDonald word list, sentiment measures were calculated based on positive, negative, and uncertain words spoken by executives and analysts.
4. Statistical Analysis: A pooled OLS regression model was employed to analyze the impact of gender on sentiment and vagueness in both management discussion (MD) and Q&A sessions.

# Key Contributions
- Enhanced Interaction Terms: The introduction of tagged interaction terms provides a new dimension for analyzing sentiment in earnings calls, which has not been extensively explored in previous literature.
- Insights into Gender Dynamics: This research contributes to understanding how gender affects communication styles in corporate settings, particularly during high-stakes interactions like earnings calls.
