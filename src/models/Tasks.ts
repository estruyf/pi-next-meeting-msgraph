interface Tasks {
  '@odata.context': string;
  '@odata.count': number;
  '@odata.nextLink': string;
  value: Task[];
}

interface Task {
  '@odata.etag': string;
  id: string;
  subject: string;
}