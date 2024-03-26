import { FormEvent, useEffect, useMemo, useState, useContext } from "react";
import { useBoolean } from "@fluentui/react-hooks"
import { Checkbox, DefaultButton, Dialog, FontIcon, Stack, Text } from "@fluentui/react";
import DOMPurify from 'dompurify';
import { AppStateContext } from '../../state/AppProvider';

import styles from "./Answer.module.css";

import { AskResponse, Feedback, historyMessageFeedback, Citation } from "../../api";
import { parseAnswer } from "./AnswerParser";

import ReactMarkdown from "react-markdown";
import remarkGfm from "remark-gfm";
import supersub from 'remark-supersub'
import { ThumbDislike20Filled, ThumbLike20Filled } from "@fluentui/react-icons";
import { XSSAllowTags } from "../../constants/xssAllowTags";

const maxRetries = 3; // Set your maximum retries limit

interface Props {
  answer: AskResponse;
  onCitationClicked?: (citation: Citation) => void;
}

export const Answer = ({ answer }: Props) => {
  const [datasheetURL, setDatasheetUrl] = useState<string>('');
  const [retryCount, setRetryCount] = useState(0);
  const [error, setError] = useState<string | null>(null);
  const [isRefAccordionOpen, { toggle: toggleIsRefAccordionOpen }] = useBoolean(false);
  const filePathTruncationLimit = 50;
  const [isFetching, setIsFetching] = useState(false);
  const initializeAnswerFeedback = (answer: AskResponse) => {
      if (answer.message_id == undefined) return undefined;
      if (answer.feedback == undefined) return undefined;
      if (answer.feedback.split(",").length > 1) return Feedback.Negative;
      if (Object.values(Feedback).includes(answer.feedback)) return answer.feedback;
      return Feedback.Neutral;
  }

  const extractAnswerWithoutCitations = (fullAnswerText: string) => {
    const citationIndex = fullAnswerText.indexOf(',"citations"');
    return citationIndex !== -1 ? fullAnswerText.substring(0, citationIndex) : fullAnswerText;
  };

  const delay = (ms: number) => new Promise((resolve) => setTimeout(resolve, ms));

  const fetchLatestAnswerStatus = async (answerId: any) => {
    try {
        // Construct the URL to your Quart microservice endpoint for checking the answer status.
        // Replace `https://yourquartmicroservice.com` with the actual URL of your Quart service.
        const url = `https://quartazurefunction.azurewebsites.net/check-answer-status/${answerId}`;

        const response = await fetch(url, {
            method: 'GET', // This endpoint is designed to be accessed with a GET request.
            headers: {
                'Content-Type': 'application/json',
                // Include any other headers your Quart service requires.
            },
        });

        if (!response.ok) {
            throw new Error(`Failed to fetch the latest answer status with status: ${response.status}`);
        }

        const data = await response.json();
        // Assuming the Quart microservice returns the updated answer directly.
        // Adjust this based on the actual structure of your response.
        return data.updatedAnswer;
    } catch (error) {
        console.error('Error fetching latest answer status:', error);
        // Decide how to handle errors: you might retry, return a default message, or take other actions.
        return "Generating answers..."; // Or handle this case differently based on your app's logic.
    }
};

  const fetchDatasheetInfo = async () => {
    if (!answer || !answer.answer.trim() || answer.answer === "Generating answers...") return;

    setIsFetching(true);
    let currentAnswer = answer.answer;
    let attempts = 0;

    // Poll for the final answer
    while (currentAnswer === "Generating answer..." && attempts < maxRetries) {
      attempts += 1;
      console.log(`Checking for final answer... Attempt ${attempts}`);
      currentAnswer = await fetchLatestAnswerStatus(answer.message_id); // This should be the updated logic for checking the answer status
      if (currentAnswer !== "Generating answer...") break; // Break the loop if we have the final answer
      await delay(4000); // Wait before checking again
    }

    if (currentAnswer === "Generating answer...") {
      console.log('Failed to obtain final answer within retry limits.');
      setIsFetching(false);
      setError('Failed to obtain final answer.');
      return;
    }

    // Continue with fetching datasheet info as before, now with the final answer
    try {
      console.log('Fetching datasheet info with final answer...');
      const payload = { chat_output: { answer: currentAnswer } };
      const response = await fetch('https://quartazurefunction.azurewebsites.net/call-apim', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify(payload),
        });

      if (!response.ok) throw new Error(`API call failed with status: ${response.status}`);
      
      const data = await response.json();
      setDatasheetUrl(data.DataSheetLink);
      setRetryCount(0);
    } catch (err: any) {
      setError(err.message || 'An unexpected error occurred');
      if (retryCount < maxRetries) setRetryCount(retryCount + 1);
    } finally {
      setIsFetching(false);
    }
  };

  useEffect(() => {
    if (answer && answer.answer.trim() && answer.answer !== "Generating answers...") {
      fetchDatasheetInfo();
    }
  }, [answer]); // Dependency array

  const chatbotResponse = answer.answer; // Use your method to parse and display the response

    // Assuming parseAnswer can handle datasheetUrl and productName
    const parsedAnswer = useMemo(() => parseAnswer(answer, datasheetURL), [answer, datasheetURL]);
    // ... (the rest of your component code)
    console.log('DataSheet URL:', datasheetURL);
    const [chevronIsExpanded, setChevronIsExpanded] = useState(isRefAccordionOpen);
    const [feedbackState, setFeedbackState] = useState(initializeAnswerFeedback(answer));
    const [isFeedbackDialogOpen, setIsFeedbackDialogOpen] = useState(false);
    const [showReportInappropriateFeedback, setShowReportInappropriateFeedback] = useState(false);
    const [negativeFeedbackList, setNegativeFeedbackList] = useState<Feedback[]>([]);
    const appStateContext = useContext(AppStateContext)
    const FEEDBACK_ENABLED = appStateContext?.state.frontendSettings?.feedback_enabled && appStateContext?.state.isCosmosDBAvailable?.cosmosDB; 
    
    const handleChevronClick = () => {
        setChevronIsExpanded(!chevronIsExpanded);
        toggleIsRefAccordionOpen();
      };

    useEffect(() => {
        setChevronIsExpanded(isRefAccordionOpen);
    }, [isRefAccordionOpen]);

    useEffect(() => {
        if (answer.message_id == undefined) return;
        
        let currentFeedbackState;
        if (appStateContext?.state.feedbackState && appStateContext?.state.feedbackState[answer.message_id]) {
            currentFeedbackState = appStateContext?.state.feedbackState[answer.message_id];
        } else {
            currentFeedbackState = initializeAnswerFeedback(answer);
        }
        setFeedbackState(currentFeedbackState)
    }, [appStateContext?.state.feedbackState, feedbackState, answer.message_id]);

    const createCitationFilepath = (citation: Citation, index: number, truncate: boolean = false) => {
        let citationFilename = "";

        if (citation.filepath && citation.chunk_id) {
            if (truncate && citation.filepath.length > filePathTruncationLimit) {
                const citationLength = citation.filepath.length;
                citationFilename = `${citation.filepath.substring(0, 20)}...${citation.filepath.substring(citationLength -20)} - Part ${parseInt(citation.chunk_id) + 1}`;
            }
            else {
                citationFilename = `${citation.filepath} - Part ${parseInt(citation.chunk_id) + 1}`;
            }
        }
        else if (citation.filepath && citation.reindex_id) {
            citationFilename = `${citation.filepath} - Part ${citation.reindex_id}`;
        }
        else {
            citationFilename = `Citation ${index}`;
        }
        return citationFilename;
    }

    const onLikeResponseClicked = async () => {
        if (answer.message_id == undefined) return;

        let newFeedbackState = feedbackState;
        // Set or unset the thumbs up state
        if (feedbackState == Feedback.Positive) {
            newFeedbackState = Feedback.Neutral;
        }
        else {
            newFeedbackState = Feedback.Positive;
        }
        appStateContext?.dispatch({ type: 'SET_FEEDBACK_STATE', payload: { answerId: answer.message_id, feedback: newFeedbackState } });
        setFeedbackState(newFeedbackState);

        // Update message feedback in db
        await historyMessageFeedback(answer.message_id, newFeedbackState);
    }

    const onDislikeResponseClicked = async () => {
        if (answer.message_id == undefined) return;

        let newFeedbackState = feedbackState;
        if (feedbackState === undefined || feedbackState === Feedback.Neutral || feedbackState === Feedback.Positive) {
            newFeedbackState = Feedback.Negative;
            setFeedbackState(newFeedbackState);
            setIsFeedbackDialogOpen(true);
        } else {
            // Reset negative feedback to neutral
            newFeedbackState = Feedback.Neutral;
            setFeedbackState(newFeedbackState);
            await historyMessageFeedback(answer.message_id, Feedback.Neutral);
        }
        appStateContext?.dispatch({ type: 'SET_FEEDBACK_STATE', payload: { answerId: answer.message_id, feedback: newFeedbackState }});
    }

    const updateFeedbackList = (ev?: FormEvent<HTMLElement | HTMLInputElement>, checked?: boolean) => {
        if (answer.message_id == undefined) return;
        let selectedFeedback = (ev?.target as HTMLInputElement)?.id as Feedback;

        let feedbackList = negativeFeedbackList.slice();
        if (checked) {
            feedbackList.push(selectedFeedback);
        } else {
            feedbackList = feedbackList.filter((f) => f !== selectedFeedback);
        }

        setNegativeFeedbackList(feedbackList);
    };

    const onSubmitNegativeFeedback = async () => {
        if (answer.message_id == undefined) return;
        await historyMessageFeedback(answer.message_id, negativeFeedbackList.join(","));
        resetFeedbackDialog();
    }

    const resetFeedbackDialog = () => {
        setIsFeedbackDialogOpen(false);
        setShowReportInappropriateFeedback(false);
        setNegativeFeedbackList([]);
    }

    const UnhelpfulFeedbackContent = () => {
        return (<>
            <div>Why wasn't this response helpful?</div>
            <Stack tokens={{childrenGap: 4}}>
                <Checkbox label="Citations are missing" id={Feedback.MissingCitation} defaultChecked={negativeFeedbackList.includes(Feedback.MissingCitation)} onChange={updateFeedbackList}></Checkbox>
                <Checkbox label="Citations are wrong" id={Feedback.WrongCitation} defaultChecked={negativeFeedbackList.includes(Feedback.WrongCitation)} onChange={updateFeedbackList}></Checkbox>
                <Checkbox label="The response is not from my data" id={Feedback.OutOfScope} defaultChecked={negativeFeedbackList.includes(Feedback.OutOfScope)} onChange={updateFeedbackList}></Checkbox>
                <Checkbox label="Inaccurate or irrelevant" id={Feedback.InaccurateOrIrrelevant} defaultChecked={negativeFeedbackList.includes(Feedback.InaccurateOrIrrelevant)} onChange={updateFeedbackList}></Checkbox>
                <Checkbox label="Other" id={Feedback.OtherUnhelpful} defaultChecked={negativeFeedbackList.includes(Feedback.OtherUnhelpful)} onChange={updateFeedbackList}></Checkbox>
            </Stack>
            <div onClick={() => setShowReportInappropriateFeedback(true)} style={{ color: "#115EA3", cursor: "pointer"}}>Report inappropriate content</div>
        </>);
    }

    const ReportInappropriateFeedbackContent = () => {
        return (
            <>
                <div>The content is <span style={{ color: "red" }} >*</span></div>
                <Stack tokens={{childrenGap: 4}}>
                    <Checkbox label="Hate speech, stereotyping, demeaning" id={Feedback.HateSpeech} defaultChecked={negativeFeedbackList.includes(Feedback.HateSpeech)} onChange={updateFeedbackList}></Checkbox>
                    <Checkbox label="Violent: glorification of violence, self-harm" id={Feedback.Violent} defaultChecked={negativeFeedbackList.includes(Feedback.Violent)} onChange={updateFeedbackList}></Checkbox>
                    <Checkbox label="Sexual: explicit content, grooming" id={Feedback.Sexual} defaultChecked={negativeFeedbackList.includes(Feedback.Sexual)} onChange={updateFeedbackList}></Checkbox>
                    <Checkbox label="Manipulative: devious, emotional, pushy, bullying" defaultChecked={negativeFeedbackList.includes(Feedback.Manipulative)} id={Feedback.Manipulative} onChange={updateFeedbackList}></Checkbox>
                    <Checkbox label="Other" id={Feedback.OtherHarmful} defaultChecked={negativeFeedbackList.includes(Feedback.OtherHarmful)} onChange={updateFeedbackList}></Checkbox>
                </Stack>
            </>
        );
    }
    const onCitationClicked = (citation: Citation, event: React.MouseEvent<HTMLElement> | React.KeyboardEvent<HTMLElement>) => {
         // Prevent the default action of the event (which is navigating to a link's href).
        event.preventDefault();

  // Here you can add whatever logic you need to handle the click.
  // For example, you could update the state to show a modal with the citation details.
  console.log(`Citation clicked: ${citation.id}`);
  // ... other logic to handle the click event
  
  return (
    <>
      {error && retryCount < maxRetries && (
        <div className="error-message">
          An error occurred: {error}
          <button onClick={() => setRetryCount((prev) => prev + 1)}>Retry</button>
        </div>
      )}
      <Stack className={styles.answerContainer} tabIndex={0}>
        {!error && (
          <ReactMarkdown
            linkTarget="_blank"
            remarkPlugins={[remarkGfm, supersub]}
            children={DOMPurify.sanitize(chatbotResponse, { ALLOWED_TAGS: XSSAllowTags })}
            className={styles.answerText}
          />
        )}
        {datasheetURL && (
          <p>
            Datasheet: <a href={datasheetURL} target="_blank" rel="noopener noreferrer">View Datasheet</a>
          </p>
        )}
        <Stack horizontal grow>
          <Stack.Item grow>
            {/* Additional content here */}
          </Stack.Item>
          <Stack.Item>
            {FEEDBACK_ENABLED && answer.message_id && (
              <Stack horizontal horizontalAlign="space-between">
                <ThumbLike20Filled
                  aria-hidden="false"
                  aria-label="Like this response"
                  onClick={onLikeResponseClicked}
                  style={feedbackState === Feedback.Positive ? { color: "darkgreen", cursor: "pointer" } : { color: "slategray", cursor: "pointer" }}
                />
                <ThumbDislike20Filled
                  aria-hidden="false"
                  aria-label="Dislike this response"
                  onClick={onDislikeResponseClicked}
                  style={feedbackState === Feedback.Negative ? { color: "darkred", cursor: "pointer" } : { color: "slategray", cursor: "pointer" }}
                />
              </Stack>
            )}
          </Stack.Item>
        </Stack>
        {!!parsedAnswer.citations.length && (
          <Stack.Item className={styles.answerFooter}>
            <Stack horizontal horizontalAlign='start' verticalAlign='center'>
              <Text
                className={styles.accordionTitle}
                onClick={toggleIsRefAccordionOpen}
                aria-label="Open references"
                tabIndex={0}
                role="button"
              >
                <span>{parsedAnswer.citations.length > 1 ? `${parsedAnswer.citations.length} references` : "1 reference"}</span>
              </Text>
              <FontIcon
                className={styles.accordionIcon}
                onClick={handleChevronClick}
                iconName={chevronIsExpanded ? 'ChevronDown' : 'ChevronRight'}
              />
            </Stack>
            {chevronIsExpanded && (
              <div style={{ marginTop: 8, display: "flex", flexFlow: "column wrap", maxHeight: "150px", gap: "4px" }}>
                {parsedAnswer.citations.map((citation, idx) => (
                  <span 
                    key={idx}
                    title={createCitationFilepath(citation, idx, true)} 
                    tabIndex={0} 
                    role="link"
                    onClick={(event) => onCitationClicked(citation, event)}
                    onKeyDown={(event) => {
                      if (event.key === "Enter" || event.key === " ") {
                        onCitationClicked(citation, event);
                      }
                    }}
                  >
                    <div className={styles.citation}>{idx + 1}</div>
                    {createCitationFilepath(citation, idx, true)}
                  </span>
                ))}
              </div>
            )}
          </Stack.Item>
        )}
      </Stack>
            <Dialog 
                onDismiss={() => {
                    resetFeedbackDialog();
                    setFeedbackState(Feedback.Neutral);
                }}
                hidden={!isFeedbackDialogOpen}
                styles={{
                    
                    main: [{
                        selectors: {
                          ['@media (min-width: 480px)']: {
                            maxWidth: '600px',
                            background: "#FFFFFF",
                            boxShadow: "0px 14px 28.8px rgba(0, 0, 0, 0.24), 0px 0px 8px rgba(0, 0, 0, 0.2)",
                            borderRadius: "8px",
                            maxHeight: '600px',
                            minHeight: '100px',
                          }
                        }
                      }]
                }}
                dialogContentProps={{
                    title: "Submit Feedback",
                    showCloseButton: true
                }}
            >
                <Stack tokens={{childrenGap: 4}}>
                    <div>Your feedback will improve this experience.</div>
                    
                    {!showReportInappropriateFeedback ? <UnhelpfulFeedbackContent/> : <ReportInappropriateFeedbackContent/>}
                    
                    <div>By pressing submit, your feedback will be visible to the application owner.</div>
                    
                    <DefaultButton disabled={negativeFeedbackList.length < 1} onClick={onSubmitNegativeFeedback}>Submit</DefaultButton>
                </Stack>

            </Dialog>
        </>
    );
};}
