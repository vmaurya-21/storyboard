import * as React from 'react';

export const Icon: React.FC<any> = ({ iconName, ...props }) => {
  return React.createElement('span', { 'data-icon-name': iconName, ...props });
};
