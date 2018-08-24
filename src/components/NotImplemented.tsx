import * as React from 'react'

// This is a quick and dirty throw away component
export default class NotImplemented extends React.PureComponent {
  public render() {
    return (
      <div className="NotImplemented centered">
        <i className="ms-Icon ms-Icon--Code" aria-hidden="true" style={{
          fontSize: "3em",
          position: "relative",
          top: "2px"
        }}></i>
        <span style={{
          fontSize: "1.5em",
          marginLeft: "12px"
        }}>
          Not implemented yet
        </span>
      </div>
    )
  }
}
