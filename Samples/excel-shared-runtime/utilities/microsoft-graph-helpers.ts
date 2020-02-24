import axios from 'axios';

export const getGraphData = async (url: string, accesstoken: string) => {
  console.log ('access token for graph data: ' + accesstoken);
    const response = await axios({
        url: url,
        method: 'get',
        headers: {'Authorization': `Bearer ${accesstoken}`}
      });
    return response;
};
