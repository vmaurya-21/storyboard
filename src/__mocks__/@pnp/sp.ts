export const spfi = jest.fn(() => ({
  web: {
    lists: {
      getByTitle: jest.fn(() => ({
        items: {
          select: jest.fn(() => ({
            get: jest.fn(() => Promise.resolve([]))
          })),
          add: jest.fn(() => Promise.resolve({ data: {} })),
          getById: jest.fn(() => ({
            update: jest.fn(() => Promise.resolve()),
            delete: jest.fn(() => Promise.resolve())
          }))
        }
      }))
    }
  }
}));

export const SPFx = jest.fn();
