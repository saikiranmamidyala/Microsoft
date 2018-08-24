import * as React from 'react'
import PropTypes from 'prop-types'

import {
  Persona,
  PersonaSize,
  PersonaPresence
} from 'office-ui-fabric-react/lib/Persona'
import {
  cdnAssetsBaseUrl
} from '../shared/SharePoint'
import {
  Callout,
  DirectionalHint
} from 'office-ui-fabric-react/lib/Callout'

import styles from './ListFaces.module.scss'

export interface IListFacesProps {
  overflow: number
  people?: IListFacesPeople[]
}

export interface IListFacesPeople {
  name: string
  imageUrl: string
  initials: string
}

export interface IListFacesState {
  renderCallout: boolean
}

export default class ListFaces extends React.Component<IListFacesProps, IListFacesState> {
  private _container: HTMLDivElement 

  constructor(props) {
    super(props)
    
    this.state = {
      renderCallout: false
    }
  }

  public render() {
    const { overflow, people } = this.props
    const { renderCallout } = this.state
    const overflowing = people.length > overflow
    const visiblePeople = overflowing ? people.slice(0, overflow) : people.slice(0)

    return (
      <div
        className={styles.listFaces}
        ref={(el) => this._container = el}
        onMouseEnter={() => this.setState({ renderCallout: true })}
        onMouseLeave={() => this.setState({ renderCallout: false })}
        >
        {visiblePeople.map((person, i) => (
          <div key={i} className={styles.listFacesPersona} style={{left:`-${i * 7}px`}}>
            <Persona
              size={PersonaSize.size24}
              presence={PersonaPresence.none}
              imageInitials={person.initials}
              imageUrl={person.imageUrl==""?`${cdnAssetsBaseUrl}/images/DefaultProfile.png` :person.imageUrl}
             
              hidePersonaDetails={true}
            />
          </div>
        ))}
        {overflowing && (
          <div className={styles.listFacesOverflow}>
            <div className={styles.listFacesOverflowCircle}>
              <span className={styles.listFacesOverflowText}>
                {people.length}
              </span>
            </div>
          </div>
        )}
        {renderCallout && (
          <Callout
            target={this._container}
            isBeakVisible={false}
            directionalHint={DirectionalHint.rightCenter}
          >
            <div className={styles.listFacesCalloutContainer}>
              {people.map((person, i) => (
                <div key={i} className={styles.listFacesCalloutPersona}>
                  <Persona
                    size={PersonaSize.large}
                    presence={PersonaPresence.none}
                    imageInitials={person.initials}
                    imageUrl={person.imageUrl==""?`${cdnAssetsBaseUrl}/images/DefaultProfile.png` :person.imageUrl}
                    
                    hidePersonaDetails={true}
                  />
                  <div className={styles.listFacesCalloutName}>
                    {person.name}
                  </div>
                </div>                
              ))}
            </div>
          </Callout>
        )}
      </div>
    )
  }
}
