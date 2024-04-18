import express, { Request, Response } from 'express';

const app = express();
const port = 3000; // Можно использовать process.env.PORT для настройки порта через переменные среды

app.get('/', (req: Request, res: Response) => {
  res.send('Hello World from Express and TypeScript!');
});

app.listen(port, () => {
  console.log(`Server running on port ${port}`);
});
