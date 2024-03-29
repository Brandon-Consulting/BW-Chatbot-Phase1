import { FormEvent, useEffect, useMemo, useState, useContext, useCallback } from "react";
import { useBoolean } from "@fluentui/react-hooks"
import { Checkbox, DefaultButton, Dialog, FontIcon, Stack, Text } from "@fluentui/react";
import DOMPurify from 'dompurify';
import { AppStateContext } from '../../state/AppProvider';
import { debounce } from 'lodash-es';

import styles from "./Answer.module.css";

import { AskResponse, Feedback, historyMessageFeedback, Citation } from "../../api";
import { parseAnswer } from "./AnswerParser";

import ReactMarkdown from "react-markdown";
import remarkGfm from "remark-gfm";
import supersub from 'remark-supersub'
import { ThumbDislike20Filled, ThumbLike20Filled } from "@fluentui/react-icons";
import { XSSAllowTags } from "../../constants/xssAllowTags";


interface Props {
    answer: AskResponse;
    // Remove onCitationClicked if it's not used
  }
  
  export const Answer = ({ answer }: Props) => {
    // Initialization and state definitions
    const [datasheetURL, setDatasheetUrl] = useState<string>('');
    const [isFetching, setIsFetching] = useState<boolean>(false);
    const [error, setError] = useState<string | null>(null);
    const [isRefAccordionOpen, { toggle: toggleIsRefAccordionOpen }] = useBoolean(false);
    const [callCount, setCallCount] = useState<number>(0);
    const maxCallAttempts = 4;
    const delay = 4500; // in milliseconds
    
    // Helper function to initialize feedback state
    const initializeAnswerFeedback = (response: AskResponse) => {
      if (!response.message_id) return undefined;
      if (!response.feedback) return undefined;
      if (response.feedback.split(',').length > 1) return Feedback.Negative;
      if (Object.values(Feedback).includes(response.feedback as Feedback)) {
        return response.feedback as Feedback;
      }
      return Feedback.Neutral;
    };
  
    // State dependent on the answer
    const [feedbackState, setFeedbackState] = useState<Feedback | undefined>(
      initializeAnswerFeedback(answer)
    );
  
    // useEffect to fetch datasheet info
    useEffect(() => {
        // Check if a call attempt should be made
        const shouldFetchData = answer && !isFetching && callCount < maxCallAttempts;
    
        // Fetch datasheet info if conditions are met
  
      const fetchData = async () => {
        setIsFetching(true);
        setError(null);
  
        try {
          const response = await fetch('https://quartazurefunction.azurewebsites.net/api/call-function', {
            method: 'POST',
            headers: {
              'Content-Type': 'application/json',
              'Cache-Control': 'no-cache',
            },
            body: JSON.stringify({
              chat_output: {
                answer: answer.answer,
              },
              request_timestamp: new Date().toISOString(),
            }),
          });
  
          if (!response.ok) {
            throw new Error(`Azure Function call failed with status: ${response.status}`);
          }
  
          const data = await response.json();
          setDatasheetUrl(data.DataSheetLink);
        } catch (e) {
          if (e instanceof Error) {
            console.error("Failed to fetch datasheet info:", e.message);
            setError(e.message);
          }
        } finally {
          setIsFetching(false);
          setCallCount(prevCount => prevCount + 1);
        }
      };
      if (shouldFetchData) {
      const timeoutId = setTimeout(fetchData, delay);
      return () => clearTimeout(timeoutId);
      }
    }, [answer, isFetching, callCount, delay]);
  
    const parsedAnswer = useMemo(() => parseAnswer(answer, datasheetURL), [answer, datasheetURL]);
    // ... (the rest of your component code)
    console.log('DataSheet URL:', datasheetURL);
    const [chevronIsExpanded, setChevronIsExpanded] = useState(isRefAccordionOpen);
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
    const filePathTruncationLimit = 50;
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

};
    return (
        <>
            <Stack className={styles.answerContainer} tabIndex={0}>
                
                <Stack.Item>
                    <Stack horizontal grow>
                        <Stack.Item grow>
                            <ReactMarkdown
                                linkTarget="_blank"
                                remarkPlugins={[remarkGfm, supersub]}
                                children={DOMPurify.sanitize(parsedAnswer.markdownFormatText, {ALLOWED_TAGS: XSSAllowTags})}
                                className={styles.answerText}
                            />
                           
                             {datasheetURL && (
                                <p>
                                    Datasheet: <a href={datasheetURL} target="_blank" rel="noopener noreferrer"></a>
                                </p>
                            )}
                            
                        </Stack.Item>
                        <Stack.Item className={styles.answerHeader}>
                            {FEEDBACK_ENABLED && answer.message_id !== undefined && <Stack horizontal horizontalAlign="space-between">
                                <ThumbLike20Filled
                                    aria-hidden="false"
                                    aria-label="Like this response"
                                    onClick={() => onLikeResponseClicked()}
                                    style={feedbackState === Feedback.Positive || appStateContext?.state.feedbackState[answer.message_id] === Feedback.Positive ? 
                                        { color: "darkgreen", cursor: "pointer" } : 
                                        { color: "slategray", cursor: "pointer" }}
                                />
                                <ThumbDislike20Filled
                                    aria-hidden="false"
                                    aria-label="Dislike this response"
                                    onClick={() => onDislikeResponseClicked()}
                                    style={(feedbackState !== Feedback.Positive && feedbackState !== Feedback.Neutral && feedbackState !== undefined) ? 
                                        { color: "darkred", cursor: "pointer" } : 
                                        { color: "slategray", cursor: "pointer" }}
                                />
                            </Stack>}
                        </Stack.Item>
                    </Stack>
                    
                </Stack.Item>
                
                <Stack horizontal className={styles.answerFooter}>
                {!!parsedAnswer.citations.length && (
                    <Stack.Item
                        onKeyDown={e => e.key === "Enter" || e.key === " " ? toggleIsRefAccordionOpen() : null}
                    >
                        <Stack style={{width: "100%"}} >
                            <Stack horizontal horizontalAlign='start' verticalAlign='center'>
                                <Text
                                    className={styles.accordionTitle}
                                    onClick={toggleIsRefAccordionOpen}
                                    aria-label="Open references"
                                    tabIndex={0}
                                    role="button"
                                >
                                <span>{parsedAnswer.citations.length > 1 ? parsedAnswer.citations.length + " references" : "1 reference"}</span>
                                </Text>
                                <FontIcon className={styles.accordionIcon}
                                onClick={handleChevronClick} iconName={chevronIsExpanded ? 'ChevronDown' : 'ChevronRight'}
                                />
                            </Stack>
                            
                        </Stack>
                                
                    </Stack.Item>
                )}
                <Stack.Item className={styles.answerDisclaimerContainer}>
                    <span className={styles.answerDisclaimer}>AI-generated content may be incorrect</span>
                </Stack.Item>
                </Stack>
                {chevronIsExpanded && 
                    <div style={{ marginTop: 8, display: "flex", flexFlow: "wrap column", maxHeight: "150px", gap: "4px" }}>
                        {parsedAnswer.citations.map((citation, idx) => {
                            return (
                                <span 
                                    title={createCitationFilepath(citation, ++idx)} 
                                    tabIndex={0} 
                                    role="link" 
                                    key={idx}
                                    className={styles.citationContainer}
                                    aria-label={createCitationFilepath(citation, idx)} 
                                    onClick={() => window.open(datasheetURL, '_blank')}
                                    onKeyDown={(event) => {
                                        if (event.key === "Enter" || event.key === " ") {
                                          window.open(datasheetURL, '_blank');
                                        }
                                      }}
                    
                                >
                                    <div className={styles.citation}>{idx}</div>
                                    {createCitationFilepath(citation, idx, true)}
                                </span>);
                        })}
                    </div>
}
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
  };
