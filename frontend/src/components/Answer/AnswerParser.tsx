import { AskResponse, Citation } from "../../api";
import { cloneDeep } from "lodash-es";

type ParsedAnswer = {
    citations: Citation[];
    markdownFormatText: string;
};

export function parseAnswer(answer: AskResponse, datasheetUrl: string): ParsedAnswer {
    let answerText = answer.answer;
    const citationLinks = answerText.match(/\[(doc\d\d?\d?)]/g);
    let updatedCitations = [] as Citation[];
    citationLinks?.forEach((link, index) => {
        // Replace citation links with datasheet URL
        let markdownLink = `[Citation ${index + 1}](${datasheetUrl})`;
        answerText = answerText.replaceAll(link, markdownLink);
    });
    {/*
    const citationLinks = answerText.match(/\[(doc\d\d?\d?)]/g);

    const lengthDocN = "[doc".length;

    let filteredCitations = [] as Citation[];
    let citationReindex = 0;
    citationLinks?.forEach(link => {
        // Replacing the links/citations with markdown link syntax
        let citationIndex = link.slice(lengthDocN, link.length - 1);
        let citation = cloneDeep(answer.citations[Number(citationIndex) - 1]) as Citation;
        if (!filteredCitations.find((c) => c.id === citationIndex) && citation) {
            // Construct markdown link for the citation
            let markdownLink = `[Citation ${citationReindex + 1}](${citation.url})`;
            answerText = answerText.replaceAll(link, markdownLink);
            citation.id = citationIndex; // original doc index to de-dupe
            citation.reindex_id = citationReindex.toString(); // reindex from 1 for display
            filteredCitations.push(citation);
            citationReindex++;
        }
    });
*/}
    // Append the datasheet hyperlink to the answer text.
    // We add two new lines before appending the hyperlink to separate it from the answer content.
    //const datasheetLink = `\n\n[Datasheet](${datasheetUrl})`;
    //answerText += datasheetLink;

    return {
        citations: [], //Pass an empty array
        markdownFormatText: answerText
    };
}
