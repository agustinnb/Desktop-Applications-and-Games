using UnityEngine;
using System.Collections;
using UnityEngine.UI;

public class NPC_Talk : MonoBehaviour
{
    [SerializeField]
    private string[] thingToSay;
    [SerializeField]
    bool repeat = false;
    private int currentSpokenStrings = 0;
    private Text TextPanel;

    void Awake()
    {
        TextPanel = GetComponentInChildren<Text>();
    }

    void OnTriggerEnter(Collider other)
    {
        if (currentSpokenStrings < thingToSay.Length)
        {
            // Speak the string.. Probably using a 'Canvas Set To World Space'
            Speak(thingToSay[currentSpokenStrings]);
            currentSpokenStrings++;
        }
        else
        {
            if (repeat)
                currentSpokenStrings = 0;
        }
    }

    void Speak(string whatToSay)
    {
        // Make TextPanel visible first, however you want to do that.. Maybe TextPanel.gameObject.SetActive(true) and then set to false after some time?
        TextPanel.text = whatToSay;
    }
}
