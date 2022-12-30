import { Configuration, OpenAIApi } from "openai";
import { apiKey } from "./apiKey";

const configuration = new Configuration({
  apiKey: apiKey
});
const openai = new OpenAIApi(configuration);

export { createProposal };

// wrap the prompt and call getCompletion
async function createProposal(source: string) {
  //const prompt = `Make the following email draft more professional: "${source}"`;
  const prompt = `Rewrite the following email draft in the style of Big Lebowski: "${source}"`;
  //const prompt =
  //  "Here is a draft email I just wrote:\n\n\"You are all a bunch of incompetent engineers. There's no way implementing this feature it's going to take a full month; I could do it in one week. I'm sure you can do better if we all stop wasting time role-playing in Scrum meetings.\"\n\nMake it polite\n\n";

  try {
    const response = await getCompletion(prompt);

    return response;
  } catch (e) {
    return "Error: " + e;
  }
}

async function getCompletion(prompt: string) {
  // const temperature = 0.7;
  // const maxTokens = 256;
  // const topP = 1;
  // const frequencyPenalty = 0;
  // const presencePenalty = 0;

  const request = {
    model: "text-davinci-003",
    prompt: prompt,
    temperature: 0.7,
    max_tokens: 256,
    top_p: 1,
    frequency_penalty: 0,
    presence_penalty: 0
  }

  const response = await openai.createCompletion(request);

  return response.data.choices[0].text;
}

// const response = await openai.createCompletion({
//   model: "text-davinci-003",
//   prompt:
//     "Here is a draft email I just wrote:\n\n\"You are all a bunch of incompetent engineers. There's no way implementing this feature it's going to take a full month; I could do it in one week. I'm sure you can do better if we all stop wasting time role-playing in Scrum meetings.\"\n\nMake it polite\n\nDear Team, \n\nI understand that implementing this feature is taking a full month, but I firmly believe that it can be done in a week if we focus our efforts and make use of our time more efficiently.  I suggest that we focus our energy on making meaningful progress instead of spending too much time in Scrum meetings. \n\nI am confident that we can make real progress if we work together and utilize our skills and expertise in the best possible way.\n\nThank you for your time and consideration. \n\nSincerely, \n[Your Name]",
//   temperature: 0.7,
//   max_tokens: 256,
//   top_p: 1,
//   frequency_penalty: 0,
//   presence_penalty: 0,
// });
