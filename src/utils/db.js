import Dexie from 'dexie';

export const db = new Dexie('myOffice');

db.version(1).stores({
  mailMergeTool: '++id, &name'
});