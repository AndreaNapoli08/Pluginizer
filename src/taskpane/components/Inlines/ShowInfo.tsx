import * as React from 'react';

export const ShowInfo = ({ info, selectedText }) => {

  return (
    <div>
        <div style={{ display: 'flex', justifyContent: 'center', marginBottom: '5px', fontSize: '20px' }}>
            Information: 
        </div>
        <div style={{ fontSize: '15px' }}>
          <b>{selectedText}: </b> {info}
        </div>
        <hr />
    </div>
  )
}