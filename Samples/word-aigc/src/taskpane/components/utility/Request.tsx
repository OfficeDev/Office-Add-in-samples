import axios, { AxiosRequestConfig } from "axios";
import qs = require("qs");

/* global console */

axios.defaults.timeout = 20000;
export const get = (url: string, config: AxiosRequestConfig = {}): Promise<any> =>
  new Promise((resolve) => {
    axios
      .get(url, {
        params: config.params,
        headers: config.headers,
        paramsSerializer: (params) => qs.stringify(params, { arrayFormat: "repeat" }),
      })
      .then((res) => {
        resolve(res);
      })
      .catch((error) => {
        console.log(error);
        errorHandler(error);
      });
  });

export const post = (url: string, data = {}, config: AxiosRequestConfig = {}): Promise<any> =>
  new Promise((resolve) => {
    axios
      .post(url, data, config)
      .then((res) => {
        resolve(res);
      })
      .catch((error) => {
        console.error(error);
        errorHandler(error);
      });
  });

const errorHandler = (error: { response: { data: any; status: number } }) => {
  console.log("@@@@@", error);
  if (error.response) {
    const data = error.response.data;
    if (error.response.status === 403) {
      console.error({
        message: "Forbidden",
        description: data.message,
      });
    }
    if (error.response.status === 401 && !(data.result && data.result.isLogin)) {
      console.error({
        message: "please login",
        description: "authentication fail",
      });
    }
  }
  return Promise.reject(error);
};
