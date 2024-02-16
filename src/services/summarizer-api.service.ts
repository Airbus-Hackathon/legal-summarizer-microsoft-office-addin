/* eslint-disable no-undef */
import axios from "axios";

export function summarize(baseURL: string, text: string) {
  const endpoint = "/summarize";
  const URL = `${baseURL}${endpoint}`;
  const data = {
    text,
  };
  return axios
    .post(URL, data, {
      headers: {
        "Content-Type": "application/json",
      },
    })
    .then(({ data }) => data);
}
