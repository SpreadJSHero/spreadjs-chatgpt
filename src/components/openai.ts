
import {Configuration, OpenAIApi} from "openai"

const configuration = new Configuration({
  apiKey: "sk-",
});
const openai = new OpenAIApi(configuration);

export { openai };